# Thinka CSV Excel Converter

一个面向 **Thinka / KNX** 配置工作的 Windows CSV 与 Excel 双向转换工具。

当前版本：**v1.4.1（PowerShell 修复版）**。本项目是独立开发的辅助工具，不代表 Thinka 官方，也不保证所有 Thinka 固件版本都支持相同的导入规则。

## 主要功能

- **CSV → Excel**：将 Thinka 导出的 UTF-8 CSV 转成可编辑的 `.xlsx`，标题自动改成中文。
- **Excel → CSV**：恢复 Thinka 原始的 11 列英文标题，移除编辑用的「配件组号」列，导出 UTF-8 无 BOM 的 CSV。
- **保护 KNX 组地址**：以文本方式保存 `6/0/1` 等组地址，避免 Excel 错误地将其转换成日期。
- **可选配件分组**：同一房间填写相同的「配件组号」，多个服务可以共用一个 UUID。未分组且缺少 UUID 的服务默认留空，由 Thinka 分配。
- **生成全部 UUID Excel**：另存一份补齐缺失 UUID 的 Excel 供后续调整；已有 UUID 不随意覆盖。
- **合规检查**：检查表头、11 列结构、配件 ID、组号冲突、服务名称及部分 KNX 组地址问题，并输出中文报告。
- **本地运行**：使用 PowerShell 5.1 与 Open XML 处理 XLSX，不依赖 Excel COM，也无需联网。

## 快速开始

1. 使用 Windows 10/11，下载仓库源码并解压。
2. 双击 **`运行自检.cmd`**，检查当前 Windows 环境中的转换、UUID 分组和合规检查功能。
3. 双击 **`启动工具.cmd`** 启动图形界面。
4. 将 Thinka 原始 CSV 拖入左侧，生成中文 XLSX；在 Excel 中修改完成后，拖入右侧转回 Thinka CSV。
5. 如需组合配件（例如一个空调配合风速服务），在相同房间内给对应服务填写相同的「配件组号」。

默认输出路径为 **当前 Windows 用户桌面下的 `Thinka文件` 文件夹**；可以在程序中另选路径。

> **重要**：项目配置涉及实际 KNX 组地址、设备名称与配件 ID。导入 Thinka 前请先备份原始 CSV，并用少量设备测试。文件合规检查只是已知规则的静态检测，不能替代 Thinka 实际导入验证。

## 关于 UUID

| 情况 | 默认处理 |
| --- | --- |
| 已有 UUID | 保留原 ID |
| 无配件组号且 UUID 为空 | 保持空白，交由 Thinka 分配 |
| 同房间、相同配件组号且 UUID 为空 | 为该组分配一个 UUID |
| 同组出现互相冲突的已有 UUID | 报错，不擅自覆盖 |

需要人工编辑每个配件 ID 时，可使用 **「生成全部 UUID Excel」** 创建独立副本。

## 文件说明

- `Thinka_Excel_Converter.ps1`：PowerShell 图形界面与 CSV/XLSX 转换源码。
- `启动工具.cmd`：启动程序。
- `运行自检.cmd`：运行本地自检。
- `使用说明.txt`：详细中文操作说明。
- `更新记录.txt`：版本更新历史。

## 兼容性与状态

- 目标平台：Windows 10/11、Windows PowerShell 5.1。
- XLSX 文件直接使用 Open XML ZIP/XML 结构读写；编辑 XLSX 时可以使用 Microsoft Excel。
- v1.4.1 修复了 PowerShell 读取 XML 时 `SelectSingleNode` 参数匹配错误。
- **尚未完成 Windows 环境和 Thinka 设备上的全面验证**。遇到问题请提供错误日志与脱敏样本，不要公开上传真实客户配置。

## 隐私和数据安全

工具在本机处理文件，不需要上传到云服务。本仓库不存放客户 CSV、XLSX、KNX 工程和设备配置文件。