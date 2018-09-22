# <a name="outlook-add-in-api-requirement-set-12"></a>Outlook 外接程序 API 要求集 1.2

Outlook 加载项 API 子集的 JavaScript API for Office 包括对象、 方法、 属性和事件，您可以在 Outlook 中使用外接程序。

> [!NOTE]
> 本文档是[要求设置](/javascript/office/requirement-sets/outlook-api-requirement-sets)而不是最新的要求集。 

## <a name="whats-new-in-12"></a>1.2 中的新增功能有哪些？

要求集 1.2 包括[要求集 1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) 的所有功能。它添加了外接程序在用户的游标中插入文本的功能，无论是在邮件的主题还是正文中皆可插入文本。

### <a name="change-log"></a>更改日志

- 添加[Office.context.mailbox.item.getSelectedDataAsync](office.context.mailbox.item.md#getselecteddataasynccoerciontype-options-callback--string)： 中的主题或邮件正文的选定的数据的异步返回。
- 添加了 [Office.context.mailbox.item.setSelectedDataAsync](office.context.mailbox.item.md#setselecteddataasyncdata-options-callback)：以异步方式将数据插入到邮件的正文或主题中。
- 修改了 [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#displayreplyallformformdata)：将 `attachments` 属性添加到 `formData` 参数。
- 修改了 [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#displayreplyformformdata)：将 `attachments` 属性添加到 `formData` 参数。

## <a name="see-also"></a>另请参阅

- [Outlook 外接程序](https://docs.microsoft.com/outlook/add-ins/)
- [Outlook 外接程序代码示例](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [入门](https://docs.microsoft.com/outlook/add-ins/quick-start)