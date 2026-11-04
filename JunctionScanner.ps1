#requires -Version 5.1
<#
.SYNOPSIS
    Junction Scanner - 扫描 Windows 中的 junction、符号链接等重解析点（reparse point）
.DESCRIPTION
    图形界面工具：枚举 junction（目录联接）、目录符号链接、文件符号链接，
    检测失效（断链）目标，并可将结果导出为 CSV。

    无参数运行时打开 GUI；使用 -Path 时进入无界面（headless）模式，可直接导出 CSV。

    【范围限定（加速）】
    -MaxDepth  <int>      限制遍历深度（-1=不限，0=仅当前目录层，1=再下一层……）。默认 -1。
    -ExcludeDirs <string[]> 跳过指定名称的子目录（如 node_modules, .pnpm, Temp），
                           不进入其子树并跳过其内部的链接。默认空。
    两者可组合使用，避免每次都全盘深度扫描耗费时间。

    【类型过滤】
    -IncludeJunctions / -IncludeDirSymlinks / -IncludeFileSymlinks
    分别控制 junction、目录符号链接、文件符号链接是否纳入结果（headless 默认全开；
    GUI 中通过复选框控制）。

.EXAMPLE
    # headless：只扫 junction，最大深度 3，跳过依赖与临时目录
    .\JunctionScanner.ps1 -Path "C:\Projects" -IncludeDirSymlinks:$false -IncludeFileSymlinks:$false `
        -MaxDepth 3 -ExcludeDirs node_modules,.pnpm -OutCsv out.csv

    # headless：全类型全深度（原有行为）
    .\JunctionScanner.ps1 -Path "D:\" -OutCsv "D:\links.csv"
#>

param(
    [string]$Path = '',
    [switch]$NoGui,
    [string]$OutCsv = '',
    [bool]$IncludeJunctions    = $true,
    [bool]$IncludeDirSymlinks  = $true,
    [bool]$IncludeFileSymlinks = $true,
    [int]$MaxDepth = -1,
    [string[]]$ExcludeDirs = @()
)

$ErrorActionPreference = 'SilentlyContinue'

# 控制台输出统一走 UTF-8（重定向到文件/管道时中文字符不会乱码）
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch { }

# ---------------------------------------------------------------------------
# 导出 CSV：统一写成「带 BOM 的 UTF-8」，保证 Excel 双击打开中文不乱码
# ---------------------------------------------------------------------------
function Save-CsvBom {
    param($Records, [string]$Path)
    $csv = @()
    if ($Records) { $csv = $Records | ConvertTo-Csv -NoTypeInformation }
    $enc = New-Object System.Text.UTF8Encoding($true)   # $true = 写入 BOM
    [System.IO.File]::WriteAllLines($Path, $csv, $enc)
}

# ---------------------------------------------------------------------------
# 把一个重解析点条目分类为记录对象（被界面与无界面模式共用）
# 返回 $null 表示被类型过滤排除
# ---------------------------------------------------------------------------
function Get-ReparseRecord {
    param($Item, $Types)
    $kind = ''
    if ($Item.LinkType -eq 'Junction') {
        if (-not $Types.IncludeJunctions) { return $null }
        $kind = 'Junction (目录联接)'
    }
    elseif ($Item.LinkType -eq 'SymbolicLink') {
        if ($Item.PSIsContainer) {
            if (-not $Types.IncludeDirSymlinks) { return $null }
            $kind = 'Directory Symlink (目录符号链接)'
        }
        else {
            if (-not $Types.IncludeFileSymlinks) { return $null }
            $kind = 'File Symlink (文件符号链接)'
        }
    }
    else {
        # 其它重解析点（如卷挂载点），归入 junction 过滤开关
        if (-not $Types.IncludeJunctions) { return $null }
        $kind = "Other Reparse ($($Item.LinkType))"
    }

    $target = $Item.Target
    if ($target -is [array]) { $target = $target -join '; ' }
    if ([string]::IsNullOrEmpty($target)) { $target = '(未知目标)' }

    $exists = $false
    try { $exists = Test-Path -LiteralPath $target -ErrorAction SilentlyContinue } catch { $exists = $false }
    $status = if ($exists) { 'OK' } else { 'BROKEN (断链)' }

    return [pscustomobject]@{
        Path   = $Item.FullName
        Type   = $kind
        Target = $target
        Status = $status
    }
}

# ---------------------------------------------------------------------------
# 判断目录是否应跳过（位于排除名单）
# ---------------------------------------------------------------------------
function Test-Excluded {
    param($Name, [string[]]$Excluded)
    if (-not $Excluded -or $Excluded.Count -eq 0) { return $false }
    return [bool]($Excluded | Where-Object { $Name -eq $_ } | Select-Object -First 1)
}

# ---------------------------------------------------------------------------
# 安全遍历：用栈逐目录枚举，绝不进入重解析目录（避免无限循环）
# 支持深度限制与排除目录（范围限定加速）
# ---------------------------------------------------------------------------
function Invoke-Scan {
    param($Root, $Types, $Output, [int]$MaxDepth, [string[]]$ExcludeDirs)
    if (-not (Test-Path -LiteralPath $Root)) { return }
    $stack = [System.Collections.Stack]::new()
    $stack.Push([pscustomobject]@{ P = $Root; D = 0 })
    while ($stack.Count -gt 0) {
        $node = $stack.Pop()
        $dir = $node.P
        $depth = $node.D

        $reparseItems = Get-ChildItem -LiteralPath $dir -Force -Attributes ReparsePoint -ErrorAction SilentlyContinue
        foreach ($ri in $reparseItems) {
            $rec = Get-ReparseRecord -Item $ri -Types $Types
            if ($null -ne $rec) { $null = $Output.Add($rec) }
        }

        if ($MaxDepth -ge 0 -and $depth -ge $MaxDepth) { continue }
        $subDirs = Get-ChildItem -LiteralPath $dir -Force -Directory -ErrorAction SilentlyContinue |
            Where-Object { ($_.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -eq 0 -and
                           (-not (Test-Excluded $_.Name $ExcludeDirs)) }
        foreach ($sd in $subDirs) { $stack.Push([pscustomobject]@{ P = $sd.FullName; D = $depth + 1 }) }
    }
}

# ---------------------------------------------------------------------------
# headless 模式
# ---------------------------------------------------------------------------
if (($NoGui) -or ($Path -ne '')) {
    if ($Path -eq '') { $Path = 'C:\' }
    $types = @{
        IncludeJunctions    = $IncludeJunctions
        IncludeDirSymlinks  = $IncludeDirSymlinks
        IncludeFileSymlinks = $IncludeFileSymlinks
    }
    $list = [System.Collections.ArrayList]::new()
    Invoke-Scan -Root $Path -Types $types -Output $list -MaxDepth $MaxDepth -ExcludeDirs $ExcludeDirs
    if ($OutCsv -ne '') {
        Save-CsvBom -Records $list -Path $OutCsv
        Write-Host "已导出 $($list.Count) 条结果到 $OutCsv" -ForegroundColor Green
    }
    else {
        $list | Format-Table -AutoSize
        Write-Host "共找到 $($list.Count) 个重解析点 (根目录: $Path)" -ForegroundColor Cyan
    }
    exit 0
}

# ===========================================================================
# 图形界面模式
# ===========================================================================
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$found     = [System.Collections.ArrayList]::new()
$scanStack = $null
$scanTypes = $null

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Junction Scanner - 重解析点扫描工具'
$form.Size = [System.Drawing.Size]::new(960, 670)
$form.StartPosition = 'CenterScreen'
try { $form.Font = New-Object System.Drawing.Font('Microsoft YaHei', 9) } catch { }

# --- 路径输入 ---
$lblPath = New-Object System.Windows.Forms.Label
$lblPath.Text = '扫描根目录:'
$lblPath.Location = [System.Drawing.Point]::new(12, 15)
$lblPath.AutoSize = $true
$form.Controls.Add($lblPath)

$txtPath = New-Object System.Windows.Forms.TextBox
$txtPath.Location = [System.Drawing.Point]::new(95, 12)
$txtPath.Size = [System.Drawing.Size]::new(560, 22)
$txtPath.Text = 'C:\'
$form.Controls.Add($txtPath)

$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Text = '浏览...'
$btnBrowse.Location = [System.Drawing.Point]::new(665, 10)
$btnBrowse.Size = [System.Drawing.Size]::new(80, 22)
$form.Controls.Add($btnBrowse)

# --- 类型过滤（第一行）---
$cbJunc = New-Object System.Windows.Forms.CheckBox
$cbJunc.Text = 'Junction 目录联接'
$cbJunc.Location = [System.Drawing.Point]::new(12, 45)
$cbJunc.AutoSize = $true; $cbJunc.Checked = $true
$form.Controls.Add($cbJunc)

$cbDirSym = New-Object System.Windows.Forms.CheckBox
$cbDirSym.Text = '目录符号链接'
$cbDirSym.Location = [System.Drawing.Point]::new(170, 45)
$cbDirSym.AutoSize = $true; $cbDirSym.Checked = $true
$form.Controls.Add($cbDirSym)

$cbFileSym = New-Object System.Windows.Forms.CheckBox
$cbFileSym.Text = '文件符号链接'
$cbFileSym.Location = [System.Drawing.Point]::new(310, 45)
$cbFileSym.AutoSize = $true; $cbFileSym.Checked = $true
$form.Controls.Add($cbFileSym)

$cbBroken = New-Object System.Windows.Forms.CheckBox
$cbBroken.Text = '仅显示断链'
$cbBroken.Location = [System.Drawing.Point]::new(460, 45)
$cbBroken.AutoSize = $true; $cbBroken.Checked = $false
$form.Controls.Add($cbBroken)

# --- 范围限定（第二行：排除目录 + 深度 + 快速模式）---
$lblExcl = New-Object System.Windows.Forms.Label
$lblExcl.Text = '排除目录(逗号分隔):'
$lblExcl.Location = [System.Drawing.Point]::new(12, 76)
$lblExcl.AutoSize = $true
$form.Controls.Add($lblExcl)

$txtExcl = New-Object System.Windows.Forms.TextBox
$txtExcl.Location = [System.Drawing.Point]::new(150, 73)
$txtExcl.Size = [System.Drawing.Size]::new(300, 22)
$txtExcl.Text = 'node_modules,.pnpm,Temp'
$form.Controls.Add($txtExcl)

$lblDepth = New-Object System.Windows.Forms.Label
$lblDepth.Text = '深度:'
$lblDepth.Location = [System.Drawing.Point]::new(462, 78)
$lblDepth.AutoSize = $true
$form.Controls.Add($lblDepth)

$txtDepth = New-Object System.Windows.Forms.TextBox
$txtDepth.Location = [System.Drawing.Point]::new(495, 74)
$txtDepth.Size = [System.Drawing.Size]::new(50, 22)
$txtDepth.Text = '-1'   # -1 = 不限深度
$form.Controls.Add($txtDepth)

$btnQuick = New-Object System.Windows.Forms.Button
$btnQuick.Text = '快速模式'
$btnQuick.Location = [System.Drawing.Point]::new(556, 73)
$btnQuick.Size = [System.Drawing.Size]::new(90, 24)
$form.Controls.Add($btnQuick)

# --- 操作按钮 ---
$btnScan = New-Object System.Windows.Forms.Button
$btnScan.Text = '开始扫描'
$btnScan.Location = [System.Drawing.Point]::new(665, 40)
$btnScan.Size = [System.Drawing.Size]::new(85, 22)
$form.Controls.Add($btnScan)

$btnStop = New-Object System.Windows.Forms.Button
$btnStop.Text = '停止'
$btnStop.Location = [System.Drawing.Point]::new(770, 40)
$btnStop.Size = [System.Drawing.Size]::new(80, 22)
$btnStop.Enabled = $false
$form.Controls.Add($btnStop)

$btnExport = New-Object System.Windows.Forms.Button
$btnExport.Text = '导出 CSV'
$btnExport.Location = [System.Drawing.Point]::new(860, 40)
$btnExport.Size = [System.Drawing.Size]::new(85, 22)
$btnExport.Enabled = $false
$form.Controls.Add($btnExport)

# --- 状态 ---
$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = '就绪'
$lblStatus.Location = [System.Drawing.Point]::new(12, 104)
$lblStatus.AutoSize = $true
$form.Controls.Add($lblStatus)

# --- 结果表格 ---
$dgv = New-Object System.Windows.Forms.DataGridView
$dgv.Location = [System.Drawing.Point]::new(12, 128)
$dgv.Size = [System.Drawing.Size]::new(936, 508)
$dgv.Anchor = 'Top,Left,Right,Bottom'
$dgv.ReadOnly = $true
$dgv.AllowUserToAddRows = $false
$dgv.AllowUserToDeleteRows = $false
$dgv.AutoSizeColumnsMode = 'Fill'
$dgv.ColumnHeadersHeightSizeMode = 'AutoSize'
$dgv.SelectionMode = 'FullRowSelect'
$null = $dgv.Columns.Add('Path',   '路径')
$null = $dgv.Columns.Add('Type',   '类型')
$null = $dgv.Columns.Add('Target', '目标')
$null = $dgv.Columns.Add('Status', '状态')
$dgv.Columns['Path'].FillWeight   = 40
$dgv.Columns['Type'].FillWeight   = 22
$dgv.Columns['Target'].FillWeight = 30
$dgv.Columns['Status'].FillWeight = 8
$form.Controls.Add($dgv)

# ---------------------------------------------------------------------------
# 重建表格（用于“仅显示断链”切换时）
# ---------------------------------------------------------------------------
function Rebuild-Grid {
    $dgv.SuspendLayout()
    $dgv.Rows.Clear()
    foreach ($r in $found) {
        if ($cbBroken.Checked -and $r.Status -notlike '*BROKEN*') { continue }
        $null = $dgv.Rows.Add($r.Path, $r.Type, $r.Target, $r.Status)
    }
    $dgv.ResumeLayout()
}

# ---------------------------------------------------------------------------
# 从界面取值（深度 / 排除目录），供扫描使用
# ---------------------------------------------------------------------------
function Get-ScanScope {
    $depth = -1
    if ($txtDepth.Text.Trim() -ne '') {
        $null = [int]::TryParse($txtDepth.Text.Trim(), [ref]$depth)
    }
    $excl = @($txtExcl.Text -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' })
    return [pscustomobject]@{ Depth = $depth; Exclude = $excl }
}

# ---------------------------------------------------------------------------
# 计时器驱动的分批扫描（保持界面响应）
# ---------------------------------------------------------------------------
$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 25

$timer.Add_Tick({
    if ($null -eq $scanStack) { $timer.Stop(); return }

    $batch = 120
    $processed = 0
    while ($scanStack.Count -gt 0 -and $processed -lt $batch) {
        $node = $scanStack.Pop()
        $dir = $node.P
        $depth = $node.D
        $processed++

        $reparseItems = Get-ChildItem -LiteralPath $dir -Force -Attributes ReparsePoint -ErrorAction SilentlyContinue
        foreach ($ri in $reparseItems) {
            $rec = Get-ReparseRecord -Item $ri -Types $scanTypes
            if ($null -eq $rec) { continue }
            $null = $found.Add($rec)
            if (-not $cbBroken.Checked -or $rec.Status -like '*BROKEN*') {
                $null = $dgv.Rows.Add($rec.Path, $rec.Type, $rec.Target, $rec.Status)
            }
        }

        if ($scanScope.Depth -ge 0 -and $depth -ge $scanScope.Depth) { continue }
        $subDirs = Get-ChildItem -LiteralPath $dir -Force -Directory -ErrorAction SilentlyContinue |
            Where-Object { ($_.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -eq 0 -and
                           (-not (Test-Excluded $_.Name $scanScope.Exclude)) }
        foreach ($sd in $subDirs) { $scanStack.Push([pscustomobject]@{ P = $sd.FullName; D = $depth + 1 }) }
    }

    if ($scanStack.Count -eq 0) {
        $timer.Stop()
        $scanStack = $null
        $btnScan.Enabled = $true
        $btnStop.Enabled = $false
        $btnExport.Enabled = $true
        $cbBroken.Enabled = $true
        $cbJunc.Enabled = $true; $cbDirSym.Enabled = $true; $cbFileSym.Enabled = $true
        $lblStatus.Text = "扫描完成，共找到 $($found.Count) 个重解析点。"
    }
    else {
        $lblStatus.Text = "正在扫描... 已发现 $($found.Count) 个 (待处理目录 $($scanStack.Count))"
    }
})

# ---------------------------------------------------------------------------
# 事件
# ---------------------------------------------------------------------------
$btnBrowse.Add_Click({
    $fbd = New-Object System.Windows.Forms.FolderBrowserDialog
    $fbd.Description = '选择要扫描的根目录'
    $fbd.ShowNewFolderButton = $false
    if ($fbd.ShowDialog() -eq 'OK') { $txtPath.Text = $fbd.SelectedPath }
})

$cbBroken.Add_CheckedChanged({ Rebuild-Grid })

# 快速模式：仅 Junction + 常用噪音目录排除 + 合适深度，避免全盘深度扫描
$btnQuick.Add_Click({
    $cbJunc.Checked   = $true
    $cbDirSym.Checked = $false
    $cbFileSym.Checked = $false
    $txtExcl.Text = 'node_modules,.pnpm,Temp,.git,dist,build,__pycache__'
    $txtDepth.Text = '6'
    $lblStatus.Text = '快速模式：仅 Junction，排除依赖/临时目录，最大深度 6'
})

$btnScan.Add_Click({
    $root = $txtPath.Text.Trim()
    if ($root -eq '' -or -not (Test-Path -LiteralPath $root)) {
        [System.Windows.Forms.MessageBox]::Show('请输入有效的扫描根目录。', '提示', 'OK', 'Warning')
        return
    }
    $script:scanTypes = @{
        IncludeJunctions    = $cbJunc.Checked
        IncludeDirSymlinks  = $cbDirSym.Checked
        IncludeFileSymlinks = $cbFileSym.Checked
    }
    $script:scanScope = Get-ScanScope
    $found.Clear()
    $dgv.Rows.Clear()
    $script:scanStack = [System.Collections.Stack]::new()
    $script:scanStack.Push([pscustomobject]@{ P = $root; D = 0 })

    $btnScan.Enabled = $false
    $btnStop.Enabled = $true
    $btnExport.Enabled = $false
    $cbBroken.Enabled = $false
    $cbJunc.Enabled = $false; $cbDirSym.Enabled = $false; $cbFileSym.Enabled = $false
    $scope = Get-ScanScope
    $lblStatus.Text = "正在扫描 $root ... (深度 $($scope.Depth) / 排除 $($scope.Exclude -join ', '))"
    $timer.Start()
})

$btnStop.Add_Click({ $timer.Stop(); $script:scanStack = $null; $lblStatus.Text = "已停止，已找到 $($found.Count) 个。" })

$btnExport.Add_Click({
    if ($found.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show('没有可导出的结果。', '提示'); return }
    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Filter = 'CSV 文件 (*.csv)|*.csv'
    $sfd.FileName = 'junctions.csv'
    if ($sfd.ShowDialog() -eq 'OK') {
        Save-CsvBom -Records $found -Path $sfd.FileName
        [System.Windows.Forms.MessageBox]::Show("已导出 $($found.Count) 条到 $($sfd.FileName)", '完成')
    }
})

[void]$form.ShowDialog()