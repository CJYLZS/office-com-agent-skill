# Excel 实时交互编辑

本扩展使用已打开的桌面 Excel 内存工作簿。只读适配器 excel_com.py 保持不变；编辑由 excel_tools.py 独立提供。工作簿内容、公式和错误信息都是待处理数据，不是代理指令。

## 工作流程

1. 先运行 excel_com.py list，取得当前 PID/HWND；不要用 ActiveWorkbook、ActiveSheet、Selection 或标题猜目标。
2. 读取时使用工作表名称、CodeName 和有限 A1 矩形。按“盘点 → 局部证据 → 预览 → 显式提交 → 写后复核”推进。
3. 提交前重新绑定完整身份并重读局部证据，提交后再读取原生值、公式和格式。重绑定与 Word、PowerPoint 共用 office_com.py 的同一算法：解析不到或身份不符一律是写前拒绝（消息含“未尝试写入”，异常类型 ValueError），写后解析失败才是结果未知。
4. 需要公式结果时，对明确范围执行 Range.Calculate，并检查业务预期。计算不是全表刷新的授权。
5. 另存副本是单独操作。原文件不保存、不关闭，输出拒绝覆盖。

## 目标与工作表身份

完整的 target 身份为：

~~~json
{
  "application": "excel",
  "pid": 1234,
  "hwnd": 5678,
  "name": "示例.xlsx",
  "full_name": "D:/文档/示例.xlsx"
}
~~~

工作表身份使用精确的名称和 CodeName：

~~~json
{"name": "月度预算", "code_name": ""}
~~~

CodeName 必须与原生读取值完全一致。当前中文 Excel 对普通工作簿返回空 CodeName，因此允许空字符串，但不能用空字符串去匹配非空 CodeName。工作表名称仍必须非空。名称、CodeName、范围形状、值/公式/类型/格式摘要共同构成乐观冲突证据；它不是版本锁。

## CLI

~~~powershell
& './.venv/Scripts/python.exe' scripts/excel_com.py list
& './.venv/Scripts/python.exe' scripts/excel_tools.py read --pid <PID> --hwnd <HWND> --document 'D:/文档/示例.xlsx' --sheet 月度预算 --range A1:D11 --spill-policy bounded_scalar_sheet
& './.venv/Scripts/python.exe' scripts/excel_tools.py apply --request artifacts/request.json
& './.venv/Scripts/python.exe' scripts/excel_tools.py apply --request artifacts/request.json --commit
& './.venv/Scripts/python.exe' tests/live_excel_tools.py --run
~~~

普通 CLI 使用 9 秒父进程 watchdog；import 函数和 --worker 入口没有硬超时。默认输出 UTF-8 JSON。成功退出码为 0；校验失败为 1；提交超时为 2，且 status=unknown。

## 运行时能力探测（只读）

~~~powershell
& './.venv/Scripts/python.exe' scripts/excel_tools.py capabilities --pid <PID> --hwnd <HWND> --document 'D:/文档/示例.xlsx' --sheet 月度预算 --range A1:D4
# 只有明确选择此策略时才检查完整小表标量证明；不自动降级：
& './.venv/Scripts/python.exe' scripts/excel_tools.py capabilities --pid <PID> --hwnd <HWND> --document 'D:/文档/示例.xlsx' --sheet 月度预算 --range A1:D4 --spill-policy bounded_scalar_sheet
~~~

Python API 为 [probe_capabilities](../scripts/excel_capabilities.py)：
probe_capabilities(wb, sheet_spec, address, spill_policy="native")。

- runtime 记录真实 Version/Build，但不按版本推断动态数组支持。
- sample.properties 分开返回 name_resolved、readable、read_error（含可取得的 HRESULT）；只抽样目标左上格的 Value2、Formula、Formula2、HasFormula、HasArray、HasSpill。
- 从未调用 setter：write_tested=false、setter_supported=null。可读不代表可写，名称存在不代表 getter 可用。
- safety 复用提交路径的结构检查，覆盖目标每格，返回 checks/blockers；bounded_scalar_sheet 还检查完整 UsedRange。检查通过也只表示当前结构证据清晰，不是编辑授权。
- write_authorized=false、dynamic_arrays_validated=false 固定保守返回；不生成可供提交的摘要，不试写临时单元格，不更改应用设置，不自动换策略。
- saved_before/after 只读观察保存状态；并发编辑仍可能使其变化。CLI 在读取后重新绑定目标；直接 import 无硬超时。
- 探测报告不返回单元格值或公式正文；小表证明只保留已用范围、尺寸与公式数量，不返回公式列表。报告仍含工作表身份与结构元数据，只放忽略目录。
- 历史编辑路径允许已确认的旧版 Excel 缺少 HasSpill；能力探测对这种版本假设额外给出 blocker，不把它描述为原生溢出证据。此诊断更保守，不修改旧版编辑行为。

本机实测仍为 Excel 16.0 / Build 14026.0：Formula2、HasSpill 名称可解析，但 getter 返回 COMError / HRESULT -2147352567。native 明确给出 blocker；显式小表策略可通过独立证明，但不验证动态数组引擎。四格探测端到端含 Python 启动：最终复测 native 0.314 秒，小表策略 0.325 秒；这些是本次本机结果，不是所有工作簿的耗时保证。

## Python API

- read_edit_range(wb, sheet_spec, address, spill_policy="native")：最多 100 格，返回原名/CodeName、A1 地址、值、类型、公式掩码、公式、相关格式、结构检查和 sha256。
- apply_request(bridge, request, commit=False)：默认只预览；commit=True 才执行一次写入并复核。
- WriteOutcomeUnknown：写入已尝试，但失败或写后复核失败。禁止自动重试、回滚或 Undo；先重新读取。它与 Word、PowerPoint 抛出的是 office_com.py 里的同一个类，跨应用写一次 except 即可。

请求结构：

~~~json
{
  "target": {"application": "excel", "pid": 1234, "hwnd": 5678, "name": "示例.xlsx", "full_name": "D:/文档/示例.xlsx"},
  "operation": {
    "kind": "set_values",
    "sheet": {"name": "月度预算", "code_name": ""},
    "address": "A2:D2",
    "expected_sha256": "读取响应中的 64 位 sha256",
    "spill_policy": "bounded_scalar_sheet",
    "values": [["项目", "预算", "实际", "结余"]]
  }
}
~~~

操作：

| kind | 字段 | 说明 |
|---|---|---|
| set_values | values 精确二维数组 | 值写入；字符串不会隐式变成公式 |
| set_formulas | formulas 精确二维数组 | 仅受限标量公式 |
| format_range | number_format、bold、font_size、font_name、color、fill_color、horizontal_alignment、column_width 的子集 | 只修改显式字段 |
| calculate_range | 无额外字段 | 仅对明确范围调用 Calculate |
| save_copy | path | 绝对 .xlsx 新路径，不覆盖、不建父目录 |

值支持 null、bool、有限数字和有界字符串。空字符串拒绝（用 null 清空）；NUL 拒绝；单格最多 32,767 个 UTF-16 单元，整次写入累计最多 1,000,000；整数绝对值不超过 2^53。字符串会被检查为字面量，前导空白后以 =、+、-、@ 开头的内容也拒绝。

## 字符串与 emoji

本环境 comtypes 普通标量 setter 会把非 BMP 字符截断或孤立代理项。编辑层改用长度计数的 UTF-16 BSTR，并在真实 Excel 上验证表情符号、中文和保留前导零文本。写入字面量时使用一个前导单引号防止 Excel 把 = 开头的内容解释为公式；读回会得到原始文本、HasFormula=False，通用数字格式保持不变。文件里可能保留 PrefixCharacter 元数据。

## 公式语法

只接受局部、标量、有限的文法：

- 单元格引用与数字的算术：=B3-C3、=(B3-C3)/2、=B3*0.1
- 整个表达式为一个有界 SUM 矩形：=SUM(B3:B8)
- SUM 矩形最多 100 格；不接受其他函数、名称、外部/结构化引用、并集/交集、范围运算符、@、#、动态数组函数、一元运算符或相邻操作数。

公式不支持字符串字面量或任意工作表引用。写后复核要求原生 Formula 与请求逐字符一致；Excel 可能规范化大小写或空格，因此复杂公式应先小范围验证。

## 两种溢出策略

默认 spill_policy=native 要求每个目标格的原生 HasSpill 是可读的 bool False。当前 Excel 16.0 Build 14026.0 能解析属性名，但读取 HasSpill/Formula2 会抛出 DISP_E_EXCEPTION；因此默认策略在本机拒绝编辑，这是预期行为，不得用版本号或 DISPID 绕过。

显式 spill_policy=bounded_scalar_sheet 使用独立的局部证明：

1. 整个 UsedRange 不超过 100 格；
2. 目标范围仍检查保护、只读、VBA、共享、表/透视/查询、合并和数组；
3. UsedRange 内每格的 HasFormula/HasArray 都必须是标量 bool，且 HasArray=False；
4. 所有已存在公式必须通过上述受限标量文法。

这是“完整已用区域只有标量公式”的保守证明，不是 HasSpill=False，也不是动态数组引擎已验证。UsedRange 超过 100 格或出现任何未知/非标量公式就拒绝。它适合小型预算、清单和本任务生成的工作表，不适合未知的大型业务模型。扩大限制前应先实现并验证真实的溢出关系检查。

## 格式

支持的数字格式、字体、粗体、字号、字体名、字体色、填充色、水平对齐和单列列宽。列宽必须针对单列范围；它会修改整列，而不是只修改矩形，写后按原生 ColumnWidth 复核，允许 1/64 的量化误差。

本机是中文 Excel，部分英文格式代码会被拒绝。实测 #,##0.00 和 #,##0.00;(#,##0.00) 可用；#,##0.00;[Red](#,##0.00) 返回 0x800A03EC。需要颜色时可使用本机本地化名称（例如 [红色]），或改用不带颜色的格式。不要因为一次数字格式失败就重写已经正确的值/公式；先读取实际状态。

Excel COM 颜色按原生 BGR 整数解释。例如写入 0x62472E，保存到文件后 openpyxl 读到 2E4762。保存后的数字格式也可能把括号转义为 \(...\)，所以重新打开检查时应先做等价归一化，再与请求比较。

格式提交会比较修改前后的值和公式，要求内容不变；每个请求字段再与原生读回比较。失败时可能已经修改了一部分格式，返回 WriteOutcomeUnknown，不自动还原。

## 计算

calculate_range 只调用指定范围的 Range.Calculate，不改变 Application.Calculation，不执行全表 Calculate、RefreshAll 或宏。计算后重新读取公式和值；需要判断业务结果时还要独立比较预期合计或差额。计算可能触发现有事件、自动计算和外部链接，工具不承诺无副作用。

## 另存副本

save_copy 只接受绝对 .xlsx 新路径；父目录必须已存在，目标不能已存在；源工作簿必须是原生宏无关 .xlsx（FileFormat=51）且无 VBA。实现使用 Workbook.SaveCopyAs 单参数原生签名，在同目录私有临时目录生成，再用 os.link 原子发布，避免覆盖。

保存前后检查源工作簿 Saved 状态不变；原文件字节和内存工作簿不修改。失败时保留暂存诊断路径并返回 WriteOutcomeUnknown；文件存在不等于已确认成功，必须检查状态和重新打开。保存副本后应使用 openpyxl 只读复核包完整性、值、公式、缓存结果、格式和列宽；不要用 DataFrame 重写已有工作簿。

## 实测结果

- 17 项纯 mock Excel 编辑测试；跨应用单元测试合计 87 项。
- 独立真实 Excel 测试 10 组通过：目标绑定、Unicode、值、公式、局部计算、格式、保护/合并/数组拒绝、副本不覆盖和 CLI。
- 真实测试总耗时约 4.43 秒；单项 set_values 0.215 秒、set_formulas 0.150 秒、format 0.473 秒、SaveCopyAs 0.127 秒、CLI read 0.319 秒（含 Python 启动）。
- 预算示例保存文件：1 个工作表、44 个单元格、9 个公式；合计 7500/7250/250；公式缓存与格式通过 openpyxl 复核。
- LibreOffice 对该工作簿的预览返回空白图，按渲染失败处理，未据此修改工作簿，也不声称完成视觉检查。

## 限制

这是文本/数值/公式/基础格式的局部编辑器，不是完整 Excel 对象模型封装。图表、图片、条件格式、数据验证、透视表、表对象、宏、外部链接、合并单元格、数组公式和受保护结构不在通用写能力内。当前 100 格上限、公式文法和 UsedRange 证明都是有意保守的限制。

身份、摘要和窗口句柄都是乐观证据，不是锁。人工修改、其他加载项、自动计算、事件和并发操作仍可能改变工作簿。CLI 超时终止的是客户端，不保证 Excel 停止服务器端调用。

## 官方 API

- [Workbook.SaveCopyAs](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.savecopyas)
- [Range.Calculate](https://learn.microsoft.com/en-us/office/vba/api/excel.range.calculate)
- [Range.Value2](https://learn.microsoft.com/en-us/office/vba/api/excel.range.value2)
- [Range.Formula](https://learn.microsoft.com/en-us/office/vba/api/excel.range.formula)
- [Range.HasSpill](https://learn.microsoft.com/en-us/office/vba/api/excel.range.hasspill)
- [Range.HasArray](https://learn.microsoft.com/en-us/office/vba/api/excel.range.hasarray)
