# PowerPoint 实时交互编辑

适用于用户明确授权的已打开演示文稿。先用 powerpoint_com.py 的 list/probe/read 重新识别目标，再使用 powerpoint_tools.py 的有限操作。连接与读取保持独立，编辑不依赖 ActivePresentation、Selection 或窗口标题。

## 从 Word 工作流借鉴的约定

采用“观察局部对象 → 生成请求 → 默认预览 → 显式提交 → 重新读取”的工作流。DSH office-docx 技能要求先检查现有内容和格式，局部修改时保留混合格式，并对成品同时检查内容与视觉；本适配器将这些约定用于 live COM。Word 原生 Range 坐标、修订及 UndoRecord 不移植到 PowerPoint。

PowerPoint 用 presentation target + SlideID + Shape.Id 定位；页码用于显示和移动目标，不能代替对象身份。文本摘要覆盖文本，不覆盖背景、格式、动画、媒体、备注或其他形状。结构操作比较当前 SlideID/Shape.Id 有序列表；它不是文档锁。

## 模板优先

先检查模板的页面尺寸、文字、字体、主题色、图片与形状几何，再通过原生 Slide.Duplicate 复制现有页面。复制保留模板页已有的布局、背景、图片和对象；随后只编辑明确文本框。需要调整字数时同时检查字号、边界与渲染结果，不能仅因 COM 调用成功就认定文字可见。

全框替换默认只用于单一格式、单段文本；中文与西文可能被 PowerPoint 分成不同 run，所以即使视觉一致也会保守拒绝。混合格式优先使用原生 Replace 或明确原生 Characters 局部范围。明确接受统一字体时，whole + preserve_strategy="first_character" 继承首字字体，可能展平内部字体差异；段落格式仍由 PowerPoint 原生继承。PowerPoint 文本坐标是原生 UTF-16 单位，emoji 后的坐标不能用 Python 字符下标推断。框架不保证所有字体继承、超链接、动画、分段格式或高级对象均无损。

新文字不做隐式归一化：\r 表示段落，\u000b 表示软换行。已有标题占位符实测可将写入的 CR 转为 VT；精确写后校验会因此报告结果未知，必须读取实际状态后继续，不能自动重复替换。需要软换行时明确传 VT。新增文本框关闭 AutoSize 并启用 WordWrap，使指定宽高保持稳定；已有对象仍保留原来的自动尺寸行为，局部文字变化可能改变其高度，必须检查读回几何和视觉。

## 写入与文件副作用

默认预览不会修改文稿或导出文件。提交参数只代表调用意图，不能替代用户授权。普通编辑不保存、关闭、退出应用，也不更改全局安全、警告或宏设置。save_copy 和 export_slide 是独立显式操作，拒绝覆盖现有路径；SaveCopyAs 保留原打开文稿的名称和路径。创建成品的任务可先 SaveCopyAs 备份用户当前未保存状态，再在授权对象上交互编辑，完成后另存成品。

文稿内容修改要求 Presentation.ReadOnly 明确可写；属性缺失、混合或未知时拒绝，文件导出/另存副本可在只读源上显式调用。每项提交前重新绑定身份、复查前置条件，提交后读取实际结果。COM 调用失败、超时或报告写入失败均可能已经产生副作用；不得自动重试、Undo 或删除新对象。先重新列举和读取，再决定后续动作。多次 COM 写入不是事务。

## 验证与性能

单元测试验证请求边界、预览、冲突与拒绝路径；opt-in live 测试只修改自己创建的演示文稿并清理自己的对象。PowerPoint 可能共用进程，禁止 Application.Quit。简单局部检查目标为 10 秒以内；库函数无硬超时，CLI 使用独立监督进程。批量创作按局部操作推进，任务整体不以 10 秒为限。

成品需检查结构、页数、标题顺序、可编辑对象与实际渲染。实时页面可用原生 Slide.Export 留下当前内存状态的视觉证据；保存副本后运行共享 Office 检查器核验 ZIP/XML 和内部关系。高级对象、动画、放映和外部链接需另做验证。

## Python 与 CLI 接口

在项目 COM 环境中导入 scripts/powerpoint_tools.py。read_shape(presentation, slide_id, shape_id) 返回原文、SHA256、原生坐标、字体 runs、段落和几何证据；apply_request(bridge, request, commit=False) 先预览，commit=True 才执行。resolve_slide / resolve_shape 按原生 ID 有界定位，不创建应用或文稿。

~~~python
from powerpoint_com import PowerPointBridge
from powerpoint_tools import read_shape, apply_request

bridge = PowerPointBridge()
# pid/hwnd 来自本次 list；slide_id/shape_id 来自本次 read。
presentation, target = bridge.select(pid, hwnd)
evidence = read_shape(presentation, slide_id, shape_id)
request = {"target": target, "operation": {
    "kind": "replace_text", "slide_id": slide_id, "shape_id": shape_id,
    "expected": evidence["text"], "text_sha256": evidence["text_sha256"],
    "strategy": "native_replace", "find": "原有短语", "text": "新的短语"}}
preview = apply_request(bridge, request)
# 用户已授权该修改，检查预览后显式提交：
result = apply_request(bridge, request, commit=True)
~~~

| kind | 必需证据与参数 |
|---|---|
| replace_text | slide_id、shape_id、expected 全框原文、text_sha256；text 新文字；native_replace 另需唯一 find；characters 另需 native_start/native_length/span_expected；whole 可显式 first_character |
| format_text | 同样的文本证据；font 子集 name/name_far_east/size/bold/italic/color，geometry 子集 left/top/width/height，或 alignment=1/2/3/4 |
| duplicate_slide | slide_id、当前完整有序 expected_slide_ids、最终 move_to 页码 |
| delete_slide | slide_id、当前完整有序 expected_slide_ids |
| add_textbox | slide_id、expected_shape_ids、text、完整 geometry；font 必含 name/name_far_east/size/color，可含 bold/italic |
| export_slide | slide_id、expected_slide_ids、绝对 path（PNG/JPG）、像素 width/height |
| save_copy | expected_slide_ids、绝对 path（PPTX）；含 VBA 的源文稿拒绝转成无宏副本 |

几何单位为 point；颜色为 Office BGR 整数；新增文本框显式关闭项目符号。ID 解析上限 500 页/500 个顶层形状，局部文字上限 32,000 UTF-16 单位、256 个字体 runs。文件操作要求输出父目录已存在，先写同目录的私有暂存目录，再用 os.link 原子发布不覆盖目标，要求文件系统支持硬链接；失败保留暂存位置用于核验。

~~~powershell
& ./.venv/Scripts/python.exe scripts/powerpoint_tools.py read --target artifacts/target.json --slide-id <ID> --shape-id <ID>
& ./.venv/Scripts/python.exe scripts/powerpoint_tools.py apply --request artifacts/request.json
& ./.venv/Scripts/python.exe scripts/powerpoint_tools.py apply --request artifacts/request.json --commit
& ./.venv/Scripts/python.exe -m unittest discover -s tests -v
& ./.venv/Scripts/python.exe tests/live_powerpoint_tools.py --run
~~~

CLI 输出 JSON 到 stdout；--target 文件直接包含完整 target 对象，--request 包含 target/operation。默认 9 秒监督超时；较慢导出/另存可显式 --timeout（1..120 秒）。import 和内部 --worker 不带硬超时。超时只终止调用方，不证明 PowerPoint 服务端停止。提交可能部分完成，报告 WriteOutcomeUnknown，不自动重试；它与 Word、Excel 用的是 office_com.py 里同一个类，拒绝消息保留英文措辞并带“未尝试写入”标记。

live_powerpoint_tools.py 需要没有残留 PowerPoint 框架：PowerPoint 是单实例应用，先前的空框架会被复用而不产生新的顶层窗口，脚本现在会显式拒绝并提示先关闭，而不是在断言处失败。

## 本轮验证结果

128 项单元测试（127 通过、1 项未支持契约跳过）通过，其中 PowerPoint 编辑测试 29 项。独立真实 PowerPoint 编辑集成测试 13 组通过，覆盖另一文稿活动时的精确目标、预览不变、文本与字体保留、局部替换、混合格式/旧证据拒绝、复制移动、可编辑文本框/格式、图片导出、两个 Saved 状态下的另存副本、不覆盖、删除和其他文稿不变。完整监督 CLI 预览实测约 0.46 秒（含 Python 启动）。测试源 ReadOnly 为原生 int 0；仅关闭自有临时文稿，未 Quit。

实际模板创作另完成 6 页演示文稿，逐页读取并检查原生 PNG，保存后共享结构检查通过，并用独立库重新打开确认 44 个非空可编辑文本形状、标题顺序与无遗留占位文字。任务内容、成品、源模板、快照和临时脚本全部留在被忽略目录，不随工具发布。

真实模板还确认了 CR/VT 占位符差异与新增文本框默认 AutoSize 问题，分别通过精确读回处理及显式关闭 AutoSize 固化。静态复核发现 partial Characters 编辑误用 first_character 会影响整框字体，现仅 whole 策略允许该参数，并有拒绝与零写入回归。上述验证不覆盖所有 Office 版本、IRM、云协作或高级对象编辑。

## 原生 API 依据

- [Slide.Duplicate](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.slide.duplicate)：返回 SlideRange，副本插在源页之后，需按目标顺序 MoveTo。
- [Presentation.SaveCopyAs](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.presentation.savecopyas)：保存副本，不替换当前打开的演示文稿。
- [TextRange.Replace](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.textrange.replace)：返回匹配替换的文本范围；后续仍需读取核验。
