# Office COM 工具设计与扩展约定

本仓库面向已打开的 Windows 桌面 Office 对象，不是离线文件转换器。Word、Excel、PowerPoint 分别提供自己的领域适配器，不能把同名 Range 或撤销接口视为同一种语义。

## 四层接口

1. **连接层**：枚举顶层与 NativeOM 子窗口，绑定当前 PID/HWND；校验 NativeOM 返回的窗口确实属于目标。文档名称与路径作为附加证据，不按标题猜文档，不依赖活动对象。句柄会复用，重新连接必须重验。
2. **观察层**：先 probe 做元数据盘点，再明确选择局部范围。Word 用原生 story 坐标，Excel 用工作表名称和有限 A1 矩形，PowerPoint 用明确幻灯片和形状集合。分页返回 next_offset，超限报错或明确声明省略，不静默假装完整。
3. **操作层**：每个编辑请求只处理一项小操作；默认预览，显式提交。验证身份、原值、公式和相关结构；拒绝未知结构，不通过移除保护、展平内容或更改全局设置绕过。
4. **证据层**：结构化 JSON 包含操作、目标、耗时、范围、局限与数据摘要。摘要覆盖什么必须写清楚；没有读到的格式、外链、对象和其他区域不受摘要保护。

## 共享内核

[office_com.py](../scripts/office_com.py) 承载三个应用共用的最小底座，应用适配器只保留自己的窗口类、绑定、证据与 CLI：

- **一个异常类型**：WriteOutcomeUnknown 由 office_com 定义，Word、Excel、PowerPoint 抛出的是同一个类，调用方写一次 except 即覆盖三者。
- **一个文本摘要**：text_hash 默认按 surrogatepass 编码（Word 场景，损坏的代理项仍要被证据覆盖）；PowerPoint 传 strict=True，拒绝把损坏文本哈希成可编辑证据。
- **一套重绑定算法**：_resolve 只做原生枚举与身份解析，两种后果由不同函数决定——rebind 用于写入前，失败一律是拒绝（未尝试写入）；rebind_after_write 用于写入后，同样的失败是 WriteOutcomeUnknown，且**不会**复用写前模板，所以“未尝试写入”不会出现在已发出写入的错误里。
- **一层 Win32 底座**：OfficeBridge 只提供 EnumWindows/GetClassNameW/GetWindowThreadProcessId 与 oleacc 签名；WordBridge/ExcelBridge/PowerPointBridge 各自实现 windows/attach/select。
- **可覆盖的消息模板**：rebind 接受 refusal/mismatch 模板，rebind_after_write 接受 label/message 模板。写前模板由 Word 用中文默认值，Excel 与 PowerPoint 用英文并保留“未尝试写入”标记；写后模板三者都用各自语言的“已发出写入、可能部分完成”，都不含该标记。PowerPoint 的读证据摘要也走 text_hash(strict=True)，与提交时比较的是同一个函数。

三个适配器的 apply_request 签名仍不统一（Word 是 apply_request(doc, identity, req, commit, bridge=)，Excel/PowerPoint 是 apply_request(bridge, request, commit)），这是既有兼容面，暂不改动。

## 共同请求与结果约定

Word 既有入口保持兼容；Excel 与 PowerPoint 的新编辑层共享以下约定，领域字段分别定义，不强行共用 COM 对象模型。

- 请求包含完整 target 与单项 operation；默认预览，commit 必须显式为 true。
- target 包含 application、pid、hwnd、name、full_name，提交前后重新绑定。Word 历史请求保留原有身份字段。
- 局部证据使用精确原生内容和稳定对象标识。Excel 是工作表名称/CodeName 与 A1 矩形；PowerPoint 是 SlideID/Shape.Id。摘要的覆盖字段在各自指南中定义。
- 预览返回 committed=false；成功提交必须经过写后复核，返回 committed=true。提交不隐含保存原文件。
- 校验阶段拒绝表示本工具尚未尝试写入；WriteOutcomeUnknown 表示已尝试写入，结果可能部分完成。CLI status=unknown 或提交超时后，先读取当前状态，不能自动重试。
- 另存副本独立请求，输出拒绝覆盖；原路径、原内存身份及 Saved 状态需复核。文件生成失败保留必要诊断，不能将文件存在视为成功。

实际契约差异与未覆盖场景见 [能力与失败边界](office-safety-matrix.md)。Word 的 CLI 与传入 bridge 的 apply_request 会在读取前、提交前和写后重新解析原生窗口；省略 bridge 的旧调用仍只比较身份副本与正文摘要，不能视为同等证据强度。部分写入现在抛出 WriteOutcomeUnknown（RuntimeError 子类），CLI 报告 status=unknown。

## 性能与故障

- 一轮简单检查以 10 秒内返回为目标；只读 CLI 使用 9 秒工作进程 watchdog，函数导入不自带硬超时。
- COM 属性调用本身跨进程；优先有界批量传输，禁止先遍历所有单元格、段落或全部幻灯片再在本地截断。
- 文本预算是在一次有界传输后核算，不等于进程内存硬限制；单个对象仍可能很大。
- 不自动重试 COM 写入。调用失败、超时甚至工作进程被终止，都可能已有副作用。报告结果未知，重新读取后交由用户判断，不自动 Undo。
- 成功输出可原子替换 JSON 文件；失败后旧输出仍可能存在。消费者必须检查本次退出码、目标与请求，不能仅凭文件存在判断成功。

## 复用与测试

通用参数化工具留在 scripts/；任务专用脚本留在被忽略的 tmp/，请求与快照留在 artifacts/。不要发布用户文件、当前路径、真实业务文本或会话窗口标识。

单元测试使用构造数据或受控 mock 验证输入边界、错误类型、分页、预算和拒绝路径；opt-in live 测试验证真正的 Office COM 行为，不能用 mock 通过代替真实连接证据。只修改并清理测试自己创建的文档。Office 应用可能复用现有进程，创建 Application 不意味着拥有整个进程，不可据此 Quit。Word、Excel 和 PowerPoint 的编辑 live 测试都仅关闭自有文档，不 Quit 应用。

测试并不建立所有 Office 版本、受保护视图、协同编辑、宏、加载项或高级对象的兼容承诺。当前不建设跨应用万能编辑器，不把一次任务脚本直接包装成通用接口。
