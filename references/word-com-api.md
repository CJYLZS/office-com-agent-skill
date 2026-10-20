# Word COM 协同编辑：Microsoft 官方 API 参考

本文面向 Windows 桌面 Word 的 COM Automation。以下签名采用 Microsoft VBA 文档的对象模型表达，Win32 入口保留 C 签名。来源均为 Microsoft 官方文档；“实现建议”是基于这些接口边界给出的工程约束，不表示 Microsoft 承诺事务或协同编辑协议。研究仅使用 web_search/web_fetch，未操作浏览器或 Word。

## 1. 从窗口取得 Word 对象：AccessibleObjectFromWindow

官方来源：

- [AccessibleObjectFromWindow](https://learn.microsoft.com/en-us/windows/win32/api/oleacc/nf-oleacc-accessibleobjectfromwindow)
- [Object Identifiers](https://learn.microsoft.com/en-us/windows/win32/winauto/object-identifiers)
- [Window.Document](https://learn.microsoft.com/en-us/office/vba/api/word.window.document)

```cpp
HRESULT AccessibleObjectFromWindow(
    HWND hwnd, DWORD dwId, REFIID riid, void **ppvObject
);
```

```vb
Window.Document As Document  ' 只读
```

- 由 Oleacc.dll 导出；成功返回 S_OK，失败可能为 E_INVALIDARG、E_NOINTERFACE 或其他 COM 错误。必须检查 HRESULT 和返回指针。
- 请求原生对象模型时指定 OBJID_NATIVEOM，通常请求 IID_IDispatch。dwId 类型为 DWORD；互操作声明应采用 Windows SDK 常量并保持其 32 位位模式。HWND 是指针大小，不能无条件声明为 32 位整数。
- 官方对应表明确：**Word 的窗口类 `_WwG` 返回 `Window` 的 IDispatch**。这是最内层文档窗口或 pane，而不是顶层 Word 框架窗口；不能把结果直接当成 Application 或 Document。
- 由返回的 `Window.Document` 取得该窗口对应文档。实现建议：枚举候选内层窗口，检查文档身份并保留对象引用；后续反复读取 `ActiveDocument` 可能因用户切换窗口而改到别的文档。
- 官方提醒用户动作可能令可访问性接口失效；窗口关闭、文档切换、返回对象不支持期望接口等都应按失败处理。重新绑定后必须重新核对目标身份。

## 2. Range、story 与坐标

官方来源：

- [Range.Start](https://learn.microsoft.com/en-us/office/vba/api/word.range.start)、[Range.End](https://learn.microsoft.com/en-us/office/vba/api/word.range.end)
- [Range.SetRange](https://learn.microsoft.com/en-us/office/vba/api/word.range.setrange)、[Range.Duplicate](https://learn.microsoft.com/en-us/office/vba/api/word.range.duplicate)
- [Document.Content](https://learn.microsoft.com/en-us/office/vba/api/word.document.content)
- [Document.StoryRanges](https://learn.microsoft.com/en-us/office/vba/api/word.document.storyranges)、[Range.NextStoryRange](https://learn.microsoft.com/en-us/office/vba/api/word.range.nextstoryrange)

```vb
Range.Start As Long             ' 读写
Range.End As Long               ' 读写
Range.SetRange(Start As Long, End As Long)
Range.Duplicate As Range        ' 只读属性；返回的 Range 可修改边界
Document.Content As Range       ' 只读；正文 story
Document.StoryRanges As StoryRanges
Range.NextStoryRange As Range
```

- Start/End 相对 **所在 story 的起点**，从 0 开始；操作区间按 `[Start, End)` 理解，Start = End 为折叠插入点。它们不是磁盘文件字节偏移。
- SetRange 官方明确：所有字符均计数，包括非打印字符、未显示的隐藏字符。不能根据屏幕可见字数计算 COM 偏移。
- 设置 Start 大于 End 会同时把 End 移到 Start；设置 End 小于 Start 会把 Start 移到 End。更新两个端点时优先使用 SetRange。
- Duplicate 允许独立改变范围边界，但仍指向同一文档内容，**不是不可变文本快照**。保存整数坐标也不能防止其他编辑使其含义变化。
- **Document.Content 只表示 main document story**，官方等价式为 `Document.StoryRanges(wdMainTextStory)`。正文快照不含页眉、页脚、脚注、尾注、批注文本或文本框中的其他 story。正文中的引用或锚点不等于相应 story 的完整文本。
- 跨 story 功能应显式枚举 StoryRanges 并沿 NextStoryRange 遍历：页眉页脚链可指向下一节的同类 story；文本框链指向下一组链接文本框；正文、脚注、尾注和批注 story 的 NextStoryRange 返回 Nothing。
- 实现建议：定位信息同时记录文档身份、story 身份、Start/End、原文和上下文。相同数字在不同 story 中不表示相同位置。

### UTF-16 与 Range offsets：必须区分的层次

官方来源：

- [BSTR](https://learn.microsoft.com/en-us/previous-versions/windows/desktop/automat/bstr)
- [Surrogates and Supplementary Characters](https://learn.microsoft.com/en-us/windows/win32/intl/surrogates-and-supplementary-characters)
- [[MS-DOC] Character Position (CP)](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/a3d44e16-7d29-46f7-bb7b-d0d8a5734f83)

COM Automation 文本用 BSTR/WCHAR 表示；Windows UTF-16 中，BMP 字符一般占 1 个 16 位 code unit，补充平面字符占代理对（2 个 code units）。**面向 COM 文本的字符串偏移转换应按 UTF-16 code units 处理，不能直接用 Unicode 码点数或视觉字符数。** 例如 `A😀B` 的 UTF-16 长度为 4，B 前的字符串偏移为 3；组合字符和 emoji 序列的视觉字形数也不等于 code unit 数。Python 可用 `len(text.encode("utf-16-le")) // 2` 计算合法字符串的 UTF-16 单元数；JavaScript 的 length 和 .NET String.Length 本身按 UTF-16 单元计数。避免在代理对中间切分。

但应准确界定官方依据：Range.Start/End/SetRange 页面描述的是 Word story 的 character positions，没有在这些页面上给出“任意复杂 Word 内容的 Range 偏移 = 导出纯文本的 UTF-16 下标”的通用等式。BSTR 的编码事实本身不能证明这个等式。MS-DOC 的 CP 页面还明确把对象锚点、段落标记、表格单元格标记算作字符；该文件格式说明可帮助理解模型，但不能直接代替 COM API 的全部契约。

实现建议：普通文本映射使用 UTF-16 长度并验证；对于字段、表格、修订、隐藏文本或其他复杂内容，以 Word 实际返回的 Range.Start/End 为准。若从导出的字符串中计算位置，先验证相应 Range.Text 与期望原文一致，不能仅凭字符串索引写入。不要在定位前清除控制字符、把段落回车改成换行、trim 文本或做 Unicode 规范化后继续复用旧偏移。对 emoji、扩展汉字、组合字符、段落和表格标记建立实际 Word 版本的映射验证。

## 3. Range.Find 与 Find.Execute

官方来源：

- [Range.Find](https://learn.microsoft.com/en-us/office/vba/api/word.range.find)
- [Find.Execute（VBA）](https://learn.microsoft.com/en-us/office/vba/api/word.find.execute)
- [Find.Execute（.NET Interop）](https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word.find.execute?view=word-pia)
- [Finding and Replacing Text or Formatting](https://learn.microsoft.com/en-us/office/vba/word/concepts/customizing-word/finding-and-replacing-text-or-formatting)

```vb
Range.Find As Find  ' 只读属性；返回的 Find 可设置查找条件
Find.Execute(FindText, MatchCase, MatchWholeWord, MatchWildcards, _
    MatchSoundsLike, MatchAllWordForms, Forward, Wrap, Format, _
    ReplaceWith, Replace, MatchKashida, MatchDiacritics, _
    MatchAlefHamza, MatchControl) As Boolean
```

上述 Execute 参数为可选 Variant；.NET Interop 对应 15 个 `ref object` 参数并返回 bool。VBA 文档的参数表还列出 MatchPrefix、MatchSuffix、MatchPhrase、IgnoreSpace、IgnorePunct，但其 Syntax 行和所核实的 Interop 签名只列出上述 15 项；不要把参数表直接拼成跨版本可用的额外位置参数。附加选项需按实际 Word 对象模型支持情况设置。

- 成功匹配返回 True。通过 **Range.Find** 查找不会改变用户 Selection，但会把该 Range 重新定义为命中的范围；Selection.Find 则会改变用户选择。协同编辑应从目标范围的 Duplicate 查找，保留范围边界。
- 显式设置 Forward、Wrap、Format、MatchCase、MatchWholeWord、MatchWildcards、MatchSoundsLike、MatchAllWordForms 等条件；调用 Find.ClearFormatting 和 Find.Replacement.ClearFormatting。不要让先前查找条件影响结果。高级语言相关选项需检查实际版本与安装语言是否支持。
- 在限定范围内检索使用 `Wrap:=wdFindStop`。wdFindContinue 会继续循环，wdFindAsk 可能弹出询问；不适合无交互的精确范围操作。
- `MatchWildcards=False` 不代表任意字符串都是完全字面值：FindText 仍支持 `^p`、`^t` 等特殊代码。需要精确匹配时明确转义规则并对命中 Range.Text 再做原文相等检查。
- 枚举多个命中时，记录每次 Start/End，并从非空命中的末尾继续；若出现零长度匹配，显式推进搜索起点或停止，防止无限循环。存在多个同文命中时结合上下文判定，不能默认改第一个。
- `Replace` 决定不替换、替换一次或全部替换；协同场景批量 ReplaceAll 会扩大影响范围，宜先验证每个目标。
- 官方说明：未另作指定时，Find 替换文本继承被替换文本的格式；MatchCase=False 时，全大写命中可能把替换文本转成全大写。若输出必须精确，需考虑这个大小写变换。

## 4. Range.Text、FormattedText 与格式风险

官方来源：[Range.Text](https://learn.microsoft.com/en-us/office/vba/api/word.range.text)、[Range.FormattedText](https://learn.microsoft.com/en-us/office/vba/api/word.range.formattedtext)。

```vb
Range.Text As String          ' 读写；读取为无格式纯文本
Range.FormattedText As Range  ' 读写；包含文字与字符格式
```

设置 Text 会替换原有范围内容；它不携带原范围的格式结构。**替换可能影响格式**，尤其是范围包含多种字符格式、段落标记、字段、书签或其他结构时，不能承诺原样保留。Find 的默认格式继承也不等于任意混合格式结构无损。避免把整个 Document.Content.Text 读出、全量重写来实现局部修改。

FormattedText 用于带格式文本的复制/替换；范围含段落标记时也包含段落格式。设置它仍会替换目标文本；若意图插入，需先折叠目标范围。它同样不提供并发隔离，也不代表无损保存所有文档结构。实现时缩小修改范围并在修改后检查实际文本和必要格式。

## 5. UndoRecord：撤销分组，不是事务

官方来源：[UndoRecord](https://learn.microsoft.com/en-us/office/vba/api/word.undorecord)、[StartCustomRecord](https://learn.microsoft.com/en-us/office/vba/api/word.undorecord.startcustomrecord)、[EndCustomRecord](https://learn.microsoft.com/en-us/office/vba/api/word.undorecord.endcustomrecord)。

```vb
Set undo = Application.UndoRecord
undo.StartCustomRecord(Name)  ' Name：可选 String
undo.EndCustomRecord
```

- Name 最多保留 64 个字符，超过会截断；省略或为空时使用首个执行命令的名称。
- 官方描述 StartCustomRecord 记录其活跃期间对 **application 的所有操作**；EndCustomRecord 结束记录。不能理解为“只包含这个客户端发出的命令”。
- **多个 COM 调用组成的并发编辑不是原子事务。** UndoRecord 没有数据库式 commit/rollback、文档锁、compare-and-swap 或隔离级别。查找、核对、替换、批注之间仍可能穿插用户操作或其他客户端更改。
- 实现建议：成功开始后用 try/finally 保证配对结束，缩短记录时间；不要跨用户交互长期打开分组。已有自定义记录可能来自其他客户端，需检测并避免错误结束别人的记录。
- 某一步失败可能留下前面已经完成的编辑。不能在发现冲突后盲目调用 Undo 作为自动回滚，因为当前撤销记录可能包含其他操作。应报告实际完成项和当前状态，重新读取目标后再决定恢复。

## 6. TrackRevisions：文档级修订状态

官方来源：[Document.TrackRevisions](https://learn.microsoft.com/en-us/office/vba/api/word.document.trackrevisions)。

```vb
Document.TrackRevisions As Boolean  ' 读写
```

True 表示该文档启用更改跟踪。官方示例分别设置 TrackRevisions 和 ShowRevisions，说明“记录修订”和“显示修订”是不同状态。

实现建议：默认尊重文档原有状态；确实需要临时更改时先读取并记录原值，仅在操作范围内设置，并在清理时谨慎恢复。该属性是文档共享状态，并非每个自动化客户端私有；简单 finally 恢复也可能覆盖期间其他人的状态更改。保护/策略或其他 COM 错误导致设置失败时要停止依赖该设置的操作。不要用接受/拒绝全部修订来清理自动化痕迹。启用修订后，原文、删除标记与可见文本的关系需要重新核实，不能继续沿用编辑前的纯文本偏移。

## 7. Comments.Add：在明确范围添加批注

官方来源：[Comments.Add](https://learn.microsoft.com/en-us/office/vba/api/word.comments.add)。

```vb
Comments.Add(Range As Range, [Text As Variant]) As Comment
```

Range 必填，指定批注附着位置；Text 可选，返回新建的 Comment。官方同时给出折叠插入点与段落范围的例子。

实现建议：使用已经绑定的目标 Document.Comments，并传入其内刚核对的 Range；避免临时读取 Selection.Range 而把批注附到用户刚移动的光标。先替换再加批注时，要重新确认希望批注附着的文本与实际范围。保留返回的 Comment 引用可用于后续核验；不要认为批注会自动描述或锁定原文本。重试前检查前一次是否已经成功，避免批注重复。

## 8. 协同编辑调用顺序与能力边界

推荐流程：从 `_WwG` 取得 Window → 核对 Window.Document → 明确 story → 读取原文与上下文 → Range.Find 定位并确认唯一性 → 写入前再次核对原文、范围和目标文档 → 短时间内进行局部编辑与所需批注 → 重新读取结果并结束撤销记录。

## 9. 原生身份重绑定

窗口句柄会被系统复用，COM 对象引用也不证明该窗口仍承载原文档。apply_request 接受可选 bridge：给出时在读取前和提交前各调用一次 bridge.select(pid, hwnd, full_name)，与请求身份比对后使用**新解析出的对象**写入；写后再解析一次，失败即抛 WriteOutcomeUnknown。写前解析失败是拒绝（未尝试写入），不应报成结果未知。

省略 bridge 的旧调用只把请求里的 target 与调用方提供的 identity 副本比较，两者可能来自同一份过期快照，证据更弱。新代码应传 bridge。重绑定缩小但不消除竞争窗口：解析与写入之间，文档仍可能被关闭或替换。

失败分类只看是否已经发出写入调用：撤销记录已开但仍在写入前的拒绝保持 ValueError（未尝试写入）；写入调用之后（含写后重绑定失败、结束撤销记录失败）抛 WriteOutcomeUnknown。清理失败不会被静默吞掉，也不会掩盖已有异常。

这只是乐观冲突检测。**“再次核对后立即写入”仍不是原子比较并替换**，两步之间存在竞争窗口；应用内串行调度本客户端的命令，也不能阻止用户或其他客户端编辑。若业务要求强原子性或严格回滚，需要独立的协调/锁定协议及所有参与者配合，不能仅凭上述 COM API 宣称实现。发生目标不唯一、原文变化、窗口失效、设置失败或写入后核验失败时，停止后续依赖步骤并报告具体状态。
