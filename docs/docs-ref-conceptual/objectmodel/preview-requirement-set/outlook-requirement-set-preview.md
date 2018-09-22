# <a name="outlook-add-in-api-preview-requirement-set"></a>Outlook 外接程序 API 预览要求集

Outlook 加载项 API 子集的 JavaScript API for Office 包括对象、 方法、 属性和事件，您可以在 Outlook 中使用外接程序。

> [!NOTE]
> 本文档是[要求设置](/javascript/office/requirement-sets/outlook-api-requirement-sets)**预览**。 此要求集尚未完全实现，客户端不会准确报告对它的支持。 不应在外接程序清单中指定此要求集。 在此要求集中引入的方法和属性应在使用前单独测试其可用性。

预览要求集包括的所有[要求集 1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md)的功能。

## <a name="features-in-preview"></a>预览阶段的功能

以下功能处于预览阶段。

- [SharedProperties](/javascript/api/outlook/office.sharedproperties) -添加新的对象，代表中共享的文件夹、 日历或邮箱的约会或邮件项目的属性。
- [Event.completed](/javascript/api/office/office.addincommands.event#completed-options-) - 新的可选参数 `options`，即包含一个有效值 `allowEvent` 的字典。此值可用于取消执行事件。
- [Office.context.mailbox.item.addFileAttachmentFromBase64Async](office.context.mailbox.item.md#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) -添加附加文件从 base64 编码到邮件或约会的新方法。
- [Office.context.mailbox.item.getInitializationContextAsync](office.context.mailbox.item.md#getinitializationcontextasyncoptions-callback) - 新增了一个函数，当[可操作邮件激活](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message)加载项时，此函数返回传递的初始化数据。
- [Office.context.mailbox.item.getSharedPropertiesAsync](office.context.mailbox.item.md#getsharedpropertiesasyncoptions-callback) -添加获取代表约会或邮件项目的 sharedProperties 对象的新方法。
- [Office.context.auth.getAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) - 现已开始支持访问 `getAccessTokenAsync`，以便加载项能够[获取 Microsoft Graph API 的访问令牌](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token)。
- [Office.MailboxEnums.DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions) -添加新的位标志枚举，指定委派权限。
- [Office.EventType](/javascript/api/office/office.eventtype) -修改以支持 OfficeThemeChanged 事件添加到`OfficeThemeChanged`条目。

## <a name="see-also"></a>另请参阅

- [Outlook 外接程序](https://docs.microsoft.com/outlook/add-ins/)
- [Outlook 外接程序代码示例](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [入门](https://docs.microsoft.com/outlook/add-ins/quick-start)