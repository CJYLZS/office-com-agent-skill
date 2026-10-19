# 高效编写 Word 临时任务脚本

专用脚本是正常工作方式：技能提供连接、定位、局部检查和编辑生命周期工具；模型按任务组合它们，不必把所有业务写成通用 CLI。任务正文、PID/HWND、试验代码放 tmp/，不入库。

## 工具选择

| 工具 | 用途与边界 |
|---|---|
| word_com.WordBridge.select(pid, hwnd) | 显式绑定已打开文档；读取 identity 并和预期目标核对 |
| word_tools.read_range(doc, start, end) | 一次读取局部文本与摘要；默认最多 20,000 Word 原生单位，不读取全篇或逐段属性 |
| word_tools.find_literal(...) | 可限制 start/end、style、limit；返回原生坐标、上下文和停止原因 |
| word_tools.inspect_range(...) | 局部 XML 结构摘要与风险提示；不是完整文档结构，更不是可编辑证明 |
| word_tools.range_risks(...) | 范围内域/控件/对象、表格状态、两端是否在域代码/结果中；零长度插入同样检查 |
| word_tools.cell_text_range(cell) | 去掉单元格结束标记并核对文本；返回 Duplicate，不改变选择 |
| word_tools.assert_expected(range, expected) | 写入前核对原文；不提供并发锁 |
| word_tools.tracked_edit(doc, label) | 修订与 UndoRecord 生命周期；不自动保存、回滚或重试，不自动核对身份与业务风险 |

## 先快后细

~~~powershell
& './.venv/Scripts/python.exe' scripts/word_com.py list
& './.venv/Scripts/python.exe' scripts/word_tools.py probe --pid <PID> --hwnd <HWND>
& './.venv/Scripts/python.exe' scripts/word_tools.py find --pid <PID> --hwnd <HWND> --text '正文标题' --style -4 --limit 3
& './.venv/Scripts/python.exe' scripts/word_tools.py read --pid <PID> --hwnd <HWND> --start <START> --end <END>
& './.venv/Scripts/python.exe' scripts/word_tools.py inspect --pid <PID> --hwnd <HWND> --start <START> --end <END> --output artifacts/local.json
~~~

- 所有新 CLI 命令都只读，以 9 秒子进程 watchdog 限制阻塞；函数 import 本身没有硬超时。超时终止客户端不保证 Word 立即结束服务器端工作，不能连续重复全量请求。
- probe 不计算页数、不遍历段落、不导出全文 XML。页数统计、全域更新、重新分页不是常规检查动作。
- read/inspect 强制局部范围。XML 字节上限默认 8 MB，但只能在 COM 已返回后检查；小范围也可能携带大图片或整个域的相关结构，必须同时使用外层只读预算。
- Find 单次 COM 仍可能阻塞，3 秒预算仅在调用之间检查；limit/time_budget/out_of_scope 都不表示查找完成。读取 stop_reason，不能把已返回的几项当作全部结果。
- Word 在目录/域内可能返回早于本轮起点的命中；工具返回 word_returned_out_of_scope 并停止，避免重复命中或死循环。此时改用已确认的正文范围或标题样式，不自动跳过未知结构。
- 样式 -2/-3/-4 分别为内置一级/二级/三级标题；实际文档可能采用自定义样式。先核对样式，不能假定所有标题都规范。
- 全文纯文本一次读取有时很快，可用于本地关键词分析，但 Python 文本位置、XML 段落序号不能转换成 Word 坐标。正式定位仍用 Find/Range 并核对原文。

## 临时脚本导入与生命周期

脚本位于 tmp/<任务名>/ 时，用 __file__ 的祖先定位技能根目录，再把 scripts 加入 sys.path；不要依赖启动目录。不固定祖先层数时，通过查找 SKILL.md 与 scripts/word_com.py 判断根目录。请求和产物路径显式基于根目录。

~~~python
from word_com import WordBridge
from word_tools import find_literal, read_range, range_risks, assert_expected, tracked_edit

doc, identity = WordBridge().select(args.pid, args.hwnd)
# 先核对 identity；定位并验证唯一性、章节、上下文与外层结构。
result = find_literal(doc, args.anchor, start=args.start, end=args.end, limit=2)
# 只有 complete=True 且确认为唯一目标时，才能据此推断范围内唯一。
# 预览输出目标与计划，默认不写；仅用户明确授权且 --commit 时进入写阶段。
# 写前重新定位锚点，核对目标/邻近上下文/必要结构摘要与原文。
# with tracked_edit(doc, '具体任务'):
#     assert_expected(target_range, expected)
#     target_range.Text = replacement
# 写后重新定位并验证目标与相邻结构，报告部分完成/未知状态，默认不保存。
~~~

只修改目标最小范围。Range.Text 不是格式保留替换器，整段/整单元格赋值可能丢失混合格式、域、书签或对象；需要格式时复制 FormattedText，再在确认过的小范围处理。cell_text_range 只负责结束标记，不证明单元格适合整体替换。合并表格不能假定 Rows/Columns/Cell(r,c) 总能均匀访问。

tracked_edit 是共享文档状态的短生命周期包装，不是事务或锁。请用户暂停键入；既有修订、审阅视图和并发状态变化需要另行检查；异常后的恢复调用本身也可能失败。禁止把写入放进会超时杀死的只读 worker，禁止因异常自动重试或自动 Undo。

## 实测教训

- 850 页文档：45,925 段、986 表、3,486 个正文域。全文文本约 0.35 秒，全文 XML 在 8 秒预算内未完成；页数不能预测 XML 导出代价。
- 新工具同机实测（含连接、不含 Python 进程启动）：probe 0.219 秒；中部 800 原生单位文本 0.060 秒；首部 800 原生单位结构检查 1.278 秒。首部局部 XML 仍达约 1.13 MB，包含目录相关结构，不能把 XML 元素数当作选定文本的精确数量。
- 空合并单元格 Text 为两个结束字符，但原生 End-Start 仅 1。不能按 Python len 把 End 减 2；辅助函数将 End 减 1 后与去掉结束字符的预期文本比较，不匹配即停止。
- 范围内 Fields.Count 为零也不能证明不在域结果中；检查 Information(44/45) 的端点状态，再检查外层结构。ContentControls.Count 同样不保证覆盖外层控件。风险工具不输出 safe=True。
- 原 word_com.apply 保留全文域/控件与表格保守拦截。专用脚本经局部与外层验证可使用底层 COM，不要为了便利删除通用安全限制。

## 验证

单元测试：python -m unittest discover -s tests -v。
独立 Word 工具测试：python tests/live_tools.py --run。该测试仅修改自己创建的未保存文档，包括重复标题、Unicode 坐标、域结果及零长度插入、空合并单元格和局部 XML。不代表所有嵌套表格、纵向合并、复杂域或版式都已覆盖。
