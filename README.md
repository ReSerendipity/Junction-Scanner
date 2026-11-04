# JunctionScanner

Windows 下扫描 **重解析点（reparse point）** 的小工具：枚举 junction（目录联接）、目录/文件符号链接，检测失效（断链）目标，并支持导出 CSV。

## 文件

| 文件 | 说明 |
| --- | --- |
| `JunctionScanner.ps1` | 主脚本，含图形界面与无界面（headless）两种模式 |
| `Run-JunctionScanner.bat` | 双击启动脚本，自动以 `Bypass` 执行策略运行 `.ps1` |

## 使用方式

### 图形界面（默认）
双击 `Run-JunctionScanner.bat`，或直接运行脚本：
```powershell
.\JunctionScanner.ps1
```

界面支持：选择根目录、按类型过滤（Junction / 目录符号链接 / 文件符号链接）、仅显示断链、排除噪音目录（如 `node_modules`）、限制遍历深度、批量分批扫描不卡界面、导出 CSV（带 BOM 的 UTF-8，Excel 直接打开不乱码）。

### 无界面 / 命令行（导出 CSV）
```powershell
# 只扫 junction，最大深度 3，跳过依赖与临时目录
.\JunctionScanner.ps1 -Path "C:\Projects" -IncludeDirSymlinks:$false -IncludeFileSymlinks:$false `
    -MaxDepth 3 -ExcludeDirs node_modules,.pnpm -OutCsv out.csv

# 全类型全深度（原有行为）
.\JunctionScanner.ps1 -Path "D:\" -OutCsv "D:\links.csv"
```

## 参数

| 参数 | 默认值 | 说明 |
| --- | --- | --- |
| `-Path` | `''` | 指定扫描根目录即进入无界面模式；留空则打开 GUI |
| `-OutCsv` | `''` | 无界面模式下导出结果的 CSV 路径 |
| `-MaxDepth` | `-1` | 遍历深度（`-1`=不限，`0`=仅当前层） |
| `-ExcludeDirs` | `@()` | 跳过的子目录名（逗号分隔的数组） |
| `-IncludeJunctions` / `-IncludeDirSymlinks` / `-IncludeFileSymlinks` | `$true` | 三类重解析点是否纳入结果 |

## 安全说明
- 遍历用显式栈实现，**绝不进入重解析目录本身**，避免跟随链接导致的无限循环。
- 仅读取元数据（路径/类型/目标/是否可达），不修改任何文件或链接。
- 需 Windows PowerShell 5.1+。

## 许可
本仓库内容仅供学习与个人使用。
