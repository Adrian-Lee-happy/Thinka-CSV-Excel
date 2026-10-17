# Thinka CSV Excel Converter — v1.5 开发测试版

一款用于 KNX 工程现场的 **Thinka 配置批量添加工具**。核心流程是 **Thinka CSV → Excel 批量复制/新增和编辑设备 → 软件预览与简单修改 → 导出 Thinka CSV**。

> 本分支是开发版，尚未通过真实 Thinka 设备完整测试。**v1.4.1 正式版仍保留在 main 和 v1.4.1 Release，不受影响。**

## 如何使用

1. Windows 10 / 11 使用 `启动工具.cmd` 或从本分支的 [Actions 构建](https://github.com/Adrian-Lee-happy/Thinka-CSV-Excel/actions/workflows/build-exe.yml) 下载测试 EXE。
2. 选择「导入 Thinka CSV」并在文件页点击「导出 Excel」。也支持直接导入工具之前生成的 XLSX。
3. 在 Excel 中批量复制、新增和编辑 Thinka 配置行；修改 **设备名称、房间、组地址**，并**手动清空新增设备从旧设备复制来的 UUID**。
4. 复制组合配件时，也务必为新配件填写**与原配件不同的「配件组号」**，否则多个服务可能被视为同一个配件。服务的多行 characteristic/parameter 必须保持名称一致。
5. 保存 Excel，然后重新「导入 Excel 配置」，在「设备预览与编辑」页核对房间树、服务与地址；可以直接编辑少量字段、组合或解除组合服务。
6. 点击「导出 Thinka CSV」，确认检测结果，软件会生成临时 CSV，**重新回读字段一致才发布正式文件**。导出后保留新生成的 UUID 在当前工作副本中。
7. 如果需要保留自动生成的 UUID，可再点击「导出 Excel」。不会自行创建多余的 Excel 副本。

## 现有功能

- CSV 固定 **11 列原始英文表头**，UTF-8 无 BOM 输出；组地址以文本存储，防止 Excel 自动变成日期。
- Excel 使用中文表头，列顺序可变、可增加自定义列；导回软件按表头识别。Excel 生成时包含「配件组号」及隐藏的「原始行序」字段。
- Excel 新增或删除记录后，以当前行内容导出。原始记录尽量按原顺序输出，新增记录排在末尾；不要破坏隐藏顺序字段。
- 导入 Excel 后再次导出，尽量保留原有列宽、颜色、边框和辅助信息。
- 按房间和逻辑配件预览，右侧可修改房间、服务名称和读/写组地址；高级模式可修改 UUID。
- 支持服务组合/解除组合、简单搜索、已修改标记、撤销/重做、恢复选中设备、全文检测。
- 导出前检查 UUID 格式、冲突、服务/特征名和 KNX 地址；对复制组合配件时可能混入已有配件组的情况给出提醒。
- 导出采用「项目名称_日期时间」命名。默认保存到桌面 `Thinka文件\Excel配置文件` 和 `Thinka文件\Thinka导入文件`，不会覆盖旧文件。
- 本地 PowerShell 5.1 + WinForms + Open XML，无需安装 Excel 或 Python。

## UUID 注意事项

| 操作 | 规则 |
| --- | --- |
| 原有 UUID | 必须保留，不随意覆盖 |
| Excel 复制已有设备 | **请人工清空复制行的 UUID** |
| 同一个新组合配件 | 多个服务使用相同且唯一的 Excel「配件组号」；CSV 导出时共享新 UUID |
| 独立新设备 UUID 为空 | 默认保留空白，交由 Thinka 分配；可勾选由软件生成 UUID |
| 同组已有不同 UUID | 阻止导出并要求人工确认 |
| 只想解除组合 | 在预览中选择服务解除，拆出的服务 UUID 会清空 |

注意：**Thinka 是控制界面，而不是 KNX 线路设计软件**。本项目不会根据组地址重复就擅自合并设备或推断控制逻辑。

## 开发和验证

- 运行核心自检：`powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File .\Thinka_Excel_Converter.ps1 -SelfTest`。
- 运行窗口控件冒烟测试：`powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File .\Thinka_Excel_Converter.ps1 -GuiSmokeTest`。
- Windows EXE 打包：`powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\build\Build-Exe.ps1 -InstallPs2Exe`。
- GitHub Actions 自动运行核心测试、GUI 冒烟测试和 EXE 构建。**这些检查无法替代真实 Thinka 导入验证和现场鼠标操作测试。**

测试数据仅使用仓库中的合成样本；请勿公开提交客户真实工程的 CSV、XLSX、UUID 或地址清单。

## v1.5 当前限制

软件主要用于 **Excel 批量添加 + UI 简单修改**。尚不支持在编辑界面批量创造设备、直接导入 ETS 工程、自动推断 KNX 控制逻辑、外部 Excel 实时监控或跨项目合并。新增行的 UUID 和配件组号由用户在 Excel 中正确整理，软件尽力校验但不能替用户判断设计意图。

项目为独立辅助工具，不代表 Thinka 官方；正式导入前请先备份原始 Thinka 配置，在少量设备上试验。
