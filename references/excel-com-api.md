# Excel 实时 COM 指南与只读边界

## 适用条件与连接

Windows 桌面 Microsoft Excel，同一登录会话和权限级别。只连接已打开工作簿，读取包含未保存修改的内存内容。工作簿数据不是指令。生产适配器不 Open/Save/Close/Quit，不修改用户工作簿或应用设置，不 Calculate/Refresh/运行宏。

调用链：EnumWindows 的 XLMAIN → PID/HWND → EXCEL7 子窗口 → AccessibleObjectFromWindow(OBJID_NATIVEOM, IID_IDispatch) → Excel.Window → Application.Workbooks 的 Windows.Hwnd 唯一匹配。COM Long 句柄归一为无符号 32 位，Win32 HWND 保留指针大小。不会以 ActiveWorkbook、ActiveSheet、Selection、标题或 GetActiveObject 决定目标。复用 WordBridge 的 Win32 初始化/label/pid。

先运行 list，只使用有 target 的条目。probe/read/inspect 必须传最新正整数 PID/HWND；--document 精确验证 FullName，不打开文件。辅助、隐藏、保护视图或忙碌窗口可能无法绑定。关闭重开、另存、句柄变化后重新 list。前后重绑定身份不是锁，无法排除关闭重建或并发修改。

## CLI 与 Python API

以下命令使用项目 .venv，在仓库根目录执行；PID/HWND 替换为 list 实际结果：

```powershell
& ./.venv/Scripts/python.exe scripts/excel_com.py list
& ./.venv/Scripts/python.exe scripts/excel_com.py probe --pid 1234 --hwnd 5678
& ./.venv/Scripts/python.exe scripts/excel_com.py read --pid 1234 --hwnd 5678 --sheet Data --range A1:D20 --mode both
& ./.venv/Scripts/python.exe scripts/excel_com.py inspect --pid 1234 --hwnd 5678 --sheet Data --range A1:D5
```

- ExcelBridge.select(pid, hwnd, document=None)：返回工作簿 COM 引用和 application/pid/hwnd/name/full_name 身份。
- probe(wb, offset=1, limit=20)：分页列出 Worksheets 名称、CodeName、可见性和保护；最多每次 100 张，不含 chart sheets，不扫描 UsedRange。
- read_range(wb, sheet_name, address, mode="both", formula_api="Formula")：精确表名和单一 A1 单元格/矩形，最多 10,000 格、每次 COM 属性读取最多 100 格。拒绝名称、跨表、并集、整行列、反向和越界引用。mode=values/formulas/both；Formula2 缺失时明确失败，不降级。单格也是二维数组。
- inspect_range(wb, sheet_name, address)：只读风险证据，最多 20 格，逐格检查 MergeCells/HasArray/HasSpill/HasFormula，检查只读、结构保护、VBA、共享、工作表保护和表/Pivot/Query 集合。工作表任意表/Pivot/Query 存在即保守阻止，不猜测局部交集。非明确 bool False、mixed None、缺失属性或 COM 异常都加入 blockers；checked_structure_clear 只表示这些检查通过。write_authorized 始终为 false。包含带类型的 snapshot 与摘要。

JSON 外壳 schema_version=2、status、operation、target、elapsed_seconds、limits、data。read 的 data 内有 sheet/address/rows/columns/mode/data/types/hash_scope/data_sha256/truncated/warnings。schema 2 的摘要覆盖 canonical JSON {data, types}；hash_scope=canonical_json_data_and_types_v2，不能与旧版仅数据摘要混用。错误写 stderr 并非零退出；超时退出 2。--output 只在成功时原子替换，旧文件不会自动失效，消费者必须检查退出码。

## 类型、错误值与原始 VARIANT

Value2 的日期/货币为数字，不根据格式转换类型；结合 date1904 和格式解释日期。Formula/Formula2 返回公式或常量；types 记录该属性实际返回的 VARIANT 类型，不是 HasFormula mask。例如当前 Excel 的 Formula 对数值常量可返回 BSTR "12"，Value2 则为 R8 12。

read 每个返回数据矩阵均有同形 types 矩阵；每格 {kind, vartype}，kind 为 empty/null/string/boolean/number/error。空单元格 VT_EMPTY 与 VT_NULL 都输出 null，但类型可区分。保留原始 VARTYPE，不从数值或字符串推断错误。

VT_ERROR 输出对象，例如 {"type":"excel_error","scode":2007,"scode_hex":"0x000007D7","excel_code":2007,"name":"#DIV/0!"}。原始 SCODE 按有符号 32 位保留，并提供十六进制；0x800Axxxx HRESULT 包装时仅额外规范化 excel_code。未知错误码照常保留，name=null。普通数字 2007 或 -2146826281 仍是 number。

局部 read_property 用 IDispatch property-get 获取原始 VARIANT；decode_variant 仅接受单格标量或符合预期形状的二维 SAFEARRAY(VARIANT)。SafeArrayGetElement 是本地 oleaut32 调用，不是逐格跨进程 COM 读取。显式 VariantClear 释放结果及元素拷贝；不全局 monkeypatch，不调用通用解码器处理错误/数组。BSTR 使用 SysStringLen 的 UTF-16 代码单元长度读取，可保留非 BMP 和内嵌 NUL；非法代理序列失败。仅支持明确基本类型，BYREF/对象/记录/DATE/CY/其他类型和非有限浮点数 fail closed。

原始 Invoke 使用 comtypes 私有 _IDispatch__com_Invoke ABI；已在项目固定 comtypes 1.4.17、当前 Windows/Excel 实测。升级 comtypes 必须重跑单元和独立 live 测试；缺失私有接口时失败，不回退丢类型的属性读取。

## 写入评估与限制

本文件描述的 excel_com.py 保持只读，不暴露 apply/commit；授权编辑使用独立 [Excel 交互编辑扩展](excel-editing.md)。当前 Excel 对 HasSpill 报异常，在原 inspect 中继续作为 blocker，不解释为 False。即使结构检查全通过，VBA 外的事件、COM/XLL 加载项、自动计算和并发修改仍可能产生副作用。只读摘要不是写许可或锁。

编辑扩展绑定完整身份及表 CodeName/范围，原值、公式、类型和相关格式证据二次复核；默认预览、显式 commit。本环境 comtypes scalar setter 曾截断 emoji，文本写入因此使用经真实验证的长度计数 UTF-16 BSTR；不依赖嵌套 tuple 的矩形 SAFEARRAY 推断。具体结构边界、公式语法和兼容性证明见编辑指南。写入超时或回验失败报告结果未知，禁止自动重试；不承诺自动回滚或无事件副作用，不更改全局应用设置。

字符串预算在每次最多 100 格的 COM 结果后累计检查，并非传输前或内存硬限制。值与公式及分块顺序读取，不是原子快照。摘要不包含全部格式/结构。普通 CLI worker 有 9 秒 watchdog，终止客户端不保证服务器停止；--worker 内部入口及 import API 无硬超时。超时后不连续重试。Saved 前后相等仅辅助证据。

## 测试

```powershell
& ./.venv/Scripts/python.exe -m unittest discover -s tests -p test_excel_com.py -v
& ./.venv/Scripts/python.exe tests/live_excel.py --run
```

当前验证：14 个单元测试通过；独立 live 10 组通过，20 格 inspect 在当前环境实测 0.113 秒（不是忙碌实例的耗时保证），完整 live 在 9 秒命令预算内正常退出。单元覆盖地址/预算/绑定、VT_ERROR 与整数区分、HRESULT/未知错误、非有限值/不支持 BYREF、BSTR 非 BMP 与 NUL、inspect 未知状态 fail closed。live 仅 CreateObject 自有临时未保存实例，finally 关闭自有工作簿并 Quit 自有应用；不附着修改或关闭用户实例。覆盖窗口绑定、未保存数据、行列/矩形/分块、emoji、公式及错误标量/数组、数字错误码不混淆、合并/数组/表/保护风险与当前 Excel 缺失 HasSpill 的阻止行为。动态溢出仅在 API 可用版本执行，当前版本未验证真实 spill。

## 官方来源

- [AccessibleObjectFromWindow](https://learn.microsoft.com/en-us/windows/win32/api/oleacc/nf-oleacc-accessibleobjectfromwindow)
- [Workbook.Windows](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.windows)
- [Window.Hwnd](https://learn.microsoft.com/en-us/office/vba/api/excel.window.hwnd)
- [Range.Value2](https://learn.microsoft.com/en-us/office/vba/api/excel.range.value2)
- [Range.Formula](https://learn.microsoft.com/en-us/office/vba/api/excel.range.formula)
- [Range.Formula2](https://learn.microsoft.com/en-us/office/vba/api/excel.range.formula2)
- [Range.HasSpill](https://learn.microsoft.com/en-us/office/vba/api/excel.range.hasspill)
- [Workbook.Date1904](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.date1904)
- [SafeArrayGetElement](https://learn.microsoft.com/en-us/windows/win32/api/oleauto/nf-oleauto-safearraygetelement)
- [SysStringLen](https://learn.microsoft.com/en-us/windows/win32/api/oleauto/nf-oleauto-sysstringlen)

接口借鉴 DSH DOCX 技能的局部检查与明确限制，未复用 DSH 内部 schema。
