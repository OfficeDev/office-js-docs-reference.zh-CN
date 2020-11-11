| Class | 域 | 说明 |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[cultureInfo](/javascript/api/excel/excel.application#cultureinfo)|基于当前系统区域性设置提供信息。|
||[decimalSeparator](/javascript/api/excel/excel.application#decimalseparator)|获取用作数值的小数分隔符的字符串。|
||[thousandsSeparator](/javascript/api/excel/excel.application#thousandsseparator)|获取一个字符串，用于将数字值的小数位数与小数的左边隔开。|
||[useSystemSeparators](/javascript/api/excel/excel.application#usesystemseparators)|指定是否启用 Excel 的系统分隔符。|
|[Comment](/javascript/api/excel/excel.comment)|[提及](/javascript/api/excel/excel.comment#mentions)|获取实体 (例如，在注释中提到的人) 。|
||[richContent](/javascript/api/excel/excel.comment#richcontent)|获取丰富的注释内容 (例如，注释) 中提及。|
||[经过](/javascript/api/excel/excel.comment#resolved)|注释线程状态。|
||[updateMentions (contentWithMentions： CommentRichContent) ](/javascript/api/excel/excel.comment#updatementions-contentwithmentions-)|使用特殊格式的字符串和提及列表更新注释内容。|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[添加 (cellAddress： Range \| string，content： CommentRichContent \| String，contentType？： Excel contenttype) ](/javascript/api/excel/excel.commentcollection#add-celladdress--content--contenttype-)|使用给定单元格上的给定内容创建新批注。|
|[CommentMention](/javascript/api/excel/excel.commentmention)|[email](/javascript/api/excel/excel.commentmention#email)|Comment 中提到的实体的电子邮件地址。|
||[id](/javascript/api/excel/excel.commentmention#id)|实体的 id。|
||[name](/javascript/api/excel/excel.commentmention#name)|Comment 中提到的实体的名称。|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[提及](/javascript/api/excel/excel.commentreply#mentions)|实体 (例如，在注释中提到的人) 。|
||[经过](/javascript/api/excel/excel.commentreply#resolved)|批注答复状态。|
||[richContent](/javascript/api/excel/excel.commentreply#richcontent)|富注释内容 (例如，注释) 中提到的内容。|
||[updateMentions (contentWithMentions： CommentRichContent) ](/javascript/api/excel/excel.commentreply#updatementions-contentwithmentions-)|使用特殊格式的字符串和提及列表更新注释内容。|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[添加 (内容： CommentRichContent \| 字符串和 contenttype？： Excel contenttype) ](/javascript/api/excel/excel.commentreplycollection#add-content--contenttype-)|为批注创建批注回复。|
|[CommentRichContent](/javascript/api/excel/excel.commentrichcontent)|[提及](/javascript/api/excel/excel.commentrichcontent#mentions)|包含所有实体的数组 (例如，注释中提到的人) 。|
||[richContent](/javascript/api/excel/excel.commentrichcontent#richcontent)|指定注释的丰富内容 (例如，注释内容包含提及，第一个提到的实体的 id 属性为0，第二个提到的实体的 id 属性为 1) 。|
|[CultureInfo](/javascript/api/excel/excel.cultureinfo)|[name](/javascript/api/excel/excel.cultureinfo#name)|以 languagecode2/regioncode2 (格式获取区域性名称，例如 "zh-tw-cn" or "en-us" ) 。|
||[numberFormat](/javascript/api/excel/excel.cultureinfo#numberformat)|定义适当的区域性格式，以显示数字。|
|[NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo)|[numberDecimalSeparator](/javascript/api/excel/excel.numberformatinfo#numberdecimalseparator)|获取用作数值的小数分隔符的字符串。|
||[numberGroupSeparator](/javascript/api/excel/excel.numberformatinfo#numbergroupseparator)|获取一个字符串，用于将数字值的小数位数与小数的左边隔开。|
|[Range](/javascript/api/excel/excel.range)|[moveTo (destinationRange：范围 \| 字符串) ](/javascript/api/excel/excel.range#moveto-destinationrange-)|将单元格的值、格式和公式从当前区域移动到目标区域，替换这些单元格中的旧信息。|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[adjustIndent (数额： number) ](/javascript/api/excel/excel.rangeformat#adjustindent-amount-)|调整范围格式的缩进量。|
|[Workbook](/javascript/api/excel/excel.workbook)|[close(closeBehavior?: Excel.CloseBehavior)](/javascript/api/excel/excel.workbook#close-closebehavior-)|关闭当前工作簿。|
||[save(saveBehavior?: Excel.SaveBehavior)](/javascript/api/excel/excel.workbook#save-savebehavior-)|保存当前工作簿。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[onRowHiddenChanged](/javascript/api/excel/excel.worksheet#onrowhiddenchanged)|在特定工作表上的一个或多个行的隐藏状态更改时发生。|
|[WorksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|[address](/javascript/api/excel/excel.worksheetcalculatedeventargs#address)|完成计算的区域的地址。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onRowHiddenChanged](/javascript/api/excel/excel.worksheetcollection#onrowhiddenchanged)|在特定工作表上的一个或多个行的隐藏状态更改时发生。|
|[WorksheetRowHiddenChangedEventArgs](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs)|[address](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#address)|获取区域地址，该地址表示特定工作表上的更改区域。|
||[changeType](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#changetype)|获取表示事件触发方式的更改类型。|
||[source](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#source)|获取事件源。|
||[type](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#type)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#worksheetid)|获取其中的数据发生更改的工作表的 ID。|
