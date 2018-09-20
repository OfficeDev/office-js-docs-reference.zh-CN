# <a name="outlook-add-in-api-requirement-set-14"></a>Outlook 外接程序 API 要求集 1.4

适用于 Office 的 JavaScript API 的 Outlook 外接程序 API 子集包括您可以在 Outlook 外接程序中使用的对象、方法、属性和事件。

> [!NOTE]
> 本文档是[要求设置](/javascript/office/requirement-sets/outlook-api-requirement-sets)而不是最新的要求集。

## <a name="whats-new-in-14"></a>1.4 中的新增功能有哪些？

要求集 1.4 包括[要求集 1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) 的所有功能。它添加了对 `Office.ui` 命名空间的访问权限。

### <a name="change-log"></a>更改日志

- 添加了 [Office.context.ui.displayDialogAsync](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-)：在 Office 主机中显示一个对话框。
- 添加了 [Office.context.ui.messageParent](/javascript/api/office/office.ui#messageparent-messageobject-)：将对话框中的消息传送到其父页/开始页。
- 添加了的[Dialog](/javascript/api/office/office.dialog)对象： 时返回的对象[`displayDialogAsync`](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-)调用方法。

## <a name="see-also"></a>另请参阅

- [Outlook 外接程序](https://docs.microsoft.com/outlook/add-ins/)
- [Outlook 外接程序代码示例](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [入门](https://docs.microsoft.com/outlook/add-ins/quick-start)