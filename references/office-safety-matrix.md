# Office 能力与失败边界

这是一份验收边界，不是“所有 Office 都兼容”的声明。公共测试只用构造数据；真实文档与报告留在被忽略的 tmp/、artifacts/。测试通过不代表人工修改、事件、加载项或 COM 服务端调用被锁住。

## 共同安全回归

[test_office_safety_contract.py](../tests/test_office_safety_contract.py) 组合三套 mock，调用各应用真实 apply_request，不通过继承已有 TestCase 重复计数。

| 行为 | Word | Excel | PowerPoint |
|---|---|---|---|
| 默认预览无内容写入、不保存/关闭/Quit | 覆盖 | 覆盖 | 覆盖 |
| 过期内容证据拒绝 | 全文摘要 + 原文 | 局部摘要 | 局部摘要/原文 |
| 写入前最后一次证据变化 | 重查正文和范围；撤销记录内失败返回 WriteOutcomeUnknown | 重绑定后重读 | 重绑定后重读 |
| 请求身份字段不一致 | 与传入 identity 比较，再与原生解析结果比较 | 原生绑定比较 | 原生绑定比较 |
| 提交前原生身份重绑定 | 覆盖（CLI 与显式 bridge） | 覆盖 | 覆盖 |
| 提交采用重新绑定的对象而非缓存对象 | 覆盖 | 覆盖 | 覆盖 |
| 初始只读拒绝 | 覆盖 | 覆盖 | 覆盖 |
| 重绑定时变为只读 | 覆盖（提交前复查） | 覆盖 | 覆盖 |
| 保护拒绝 | ProtectionType 初始检查 | 工作簿/工作表，提交前复查 | 仅 ReadOnly，无独立 Protected View 证明 |
| 部分写入抛错，不重试/不 Undo | WriteOutcomeUnknown（RuntimeError 子类） | WriteOutcomeUnknown | WriteOutcomeUnknown |
| 写后身份丢失 | 结果未知 | 结果未知 | 结果未知 |

本轮单元测试 122 项：121 项通过，1 项未支持契约以显式 skip 保留，不把 skip 计为通过。其中共用契约测试 20 项（19 通过、1 跳过），Excel 能力探测测试 15 项通过。

Word 现在与另外两个应用同层：读取前、提交前和写后都会重新解析原生窗口；解析失败在写前是拒绝（未尝试写入），写后是 WriteOutcomeUnknown。撤销分组和修订设置恢复仍不是事务或回滚。省略 bridge 的旧调用保留原有成功路径与较弱证据，但失败也统一为 WriteOutcomeUnknown（RuntimeError 子类，消息文本已变），CLI 报告 status=unknown。只有真正发出写入调用后才判为 unknown：撤销记录内、写入前的最后一次拒绝仍是 ValueError。

## Excel 能力探测

[excel_capabilities.py](../scripts/excel_capabilities.py) 只读取指定目标：属性名解析与实际 getter 调用分开，setter 从未测试，不试写空白单元格。

- 目标不超过 100 格，结构证据逐格检查；getter 能力仅抽样左上格。
- native 不自动切换为小表模式；显式 bounded_scalar_sheet 才检查 <=100 格的完整 UsedRange 标量清单。
- 探测结果固定 write_authorized=false、dynamic_arrays_validated=false；无提交摘要，不替代预览与写前复查。
- 探测报告省略单元格内容与公式正文；已有编辑证据保留原有内容以支持冲突检测。
- 旧版编辑路径既有的版本例外不视为原生 HasSpill 证据；探测对此另加 blocker。
- 大表不会因探测失败而自动扩大扫描或编辑限制。大型工作表的局部安全证明仍待实现。

详见 [Excel 交互编辑指南](excel-editing.md)。

## 本轮真实 Office 验证

仅创建并清理测试自己拥有的文档，三套 live 编辑测试都不调用 Application.Quit。Word 原有两个 runner 的 Quit 已移除，避免从 CreateObject 错误推断进程所有权。

| 应用 | 本机结果 | 已测端到端局部 CLI 耗时 |
|---|---|---|
| Word | live_smoke 11 组 + live_tools 5 组通过（含写前拒绝与 CLI apply） | CLI apply 预览 0.155 秒 / 提交 0.211 秒（含 Python 启动） |
| Excel | 编辑与能力探测 11 组通过；最终复测整套约 5.46 秒 | read 0.365 秒；四格 capabilities native 0.314 秒 / 小表策略 0.325 秒 |
| PowerPoint | 13 组通过 | preview 0.506 秒 |

上述耗时包含对应 CLI Python 启动；整套测试耗时不是单次检查耗时。测试环境、Office 冷启动、类型库生成、文档结构均可能影响耗时。CLI watchdog 不保证 Excel/PowerPoint 服务端停止调用；直接 import 函数没有硬超时。10 秒是简单检查的验收目标，不是所有操作的无条件 SLA。

本机 Excel 16.0 Build 14026.0：Formula2/HasSpill 名称解析成功，但 getter 抛 COMError，HRESULT=-2147352567。未在现代动态数组 Excel 上验证，不得据此宣称兼容。

## 仍待验收

1. 真实受保护视图、Office 模态对话框、调用被拒绝、进程断开和超时：需要隔离且有清理方案的故障注入，不能拿用户文档模拟。
2. 多进程/同名文档与窗口句柄复用的跨版本真实矩阵；已有 mock 绑定测试不能替代这项验收。
3. 现代 Excel 动态数组引擎与大表小范围溢出关系证明。在证据齐全前保持现有拒绝边界，不直接放大 100 格限制。
4. 在自有 Word 文档上补齐端到端局部只读性能测试，以及冷启动/繁忙状态的单独耗时分类。
