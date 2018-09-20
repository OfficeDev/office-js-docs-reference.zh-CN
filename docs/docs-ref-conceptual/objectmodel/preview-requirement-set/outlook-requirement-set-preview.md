# <a name="outlook-add-in-api-preview-requirement-set"></a>Outlook 外接程序 API 预览要求集

适用于 Office 的 JavaScript API 的 Outlook 外接程序 API 子集包括您可以在 Outlook 外接程序中使用的对象、方法、属性和事件。

> [!NOTE]
> 本文档是[要求设置](/javascript/office/requirement-sets/outlook-api-requirement-sets)**预览**。 此要求集尚未完全实现，客户端不会准确报告对它的支持。 不应在外接程序清单中指定此要求集。 在此要求集中引入的方法和属性应在使用前单独测试其可用性。

预览要求集包括的所有[要求集 1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md)的功能。

## <a name="features-in-preview"></a>预览阶段的功能

以下功能处于预览阶段。

- [从](/javascript/api/outlook/office.from)-添加一个新对象，提供一个方法来获取值。
- [组织者](/javascript/api/outlook/office.organizer)-添加新的对象，提供要获取组织者值的方法。
- [定期](/javascript/api/outlook/office.recurrence)-添加新的对象提供方法获取和设置约会的定期模式，但仅获取定期模式会议请求的消息的数目。
- [SeriesTime](/javascript/api/outlook/office.seriestime) -添加新的对象提供获取和设置定期系列中的日期和时间的约会，并可获取的日期和时间的会议请求定期系列中的方法。
- [Event.completed](/javascript/api/office/office.addincommands.event#completed-options-) - 新的可选参数 `options`，即包含一个有效值 `allowEvent` 的字典。此值可用于取消执行事件。
- [Office.context.mailbox.item.addFileAttachmentFromBase64Async](office.context.mailbox.item.md#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) -添加附加文件从 base64 编码到邮件或约会的新方法。
- [Office.context.mailbox.item.addHandlerAsync](office.context.mailbox.item.md#addhandlerasynceventtype-handler-options-callback) -添加支持事件添加事件处理程序的新方法。
- [Office.context.mailbox.item.from](office.context.mailbox.item.md#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) -修改以获取在撰写模式下的值。
- [Office.context.mailbox.item.getInitializationContextAsync](office.context.mailbox.item.md#getinitializationcontextasyncoptions-callback) - 新增了一个函数，当[可操作邮件激活](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message)加载项时，此函数返回传递的初始化数据。
- [Office.context.mailbox.item.organizer](office.context.mailbox.item.md#organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer) -修改以在撰写模式下获取组织者值。
- [Office.context.mailbox.item.recurrence](office.context.mailbox.item.md#nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence) -添加了新的属性来获取或设置一个值，它提供用于管理定期模式的约会项目的方法。 此属性还可用于获取定期模式的会议请求项目。
- [Office.context.mailbox.item.removeHandlerAsync](office.context.mailbox.item.md#removehandlerasynceventtype-handler-options-callback) -添加移除事件处理程序的新方法。
- [Office.context.mailbox.item.seriesId](office.context.mailbox.item.md#nullable-seriesid-string) -添加了属于获取出现系列的 id 的新属性。
- [Office.context.auth.getAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) - 现已开始支持访问 `getAccessTokenAsync`，以便加载项能够[获取 Microsoft Graph API 的访问令牌](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token)。
- [Office.MailboxEnums.Days](/javascript/api/outlook/office.mailboxenums.days) -添加指定的天的周或一天中的类型的新枚举。
- [Office.MailboxEnums.Month](/javascript/api/outlook/office.mailboxenums.month) -添加指定的月份新枚举。
- [Office.MailboxEnums.RecurrenceTimeZone](/javascript/api/outlook/office.mailboxenums.recurrencetimezone) -添加指定时区定期应用于新枚举。
- [Office.MailboxEnums.RecurrenceType](/javascript/api/outlook/office.mailboxenums.recurrencetype) -添加指定类型的定期新枚举。
- [Office.MailboxEnums.WeekNumber](/javascript/api/outlook/office.mailboxenums.weeknumber) -添加指定月的周的新枚举。
- [Office.EventType](/javascript/api/office/office.eventtype) -修改以支持 RecurrenceChanged、 RecipientsChanged、 AppointmentTimeChanged 和 OfficeThemeChanged 事件通过添加`RecurrencePatternChanged`， `RecipientsChanged`， `AppointmentTimeChanged`，和`OfficeThemeChanged`条目分别。

## <a name="see-also"></a>另请参阅

- [Outlook 外接程序](https://docs.microsoft.com/outlook/add-ins/)
- [Outlook 外接程序代码示例](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [入门](https://docs.microsoft.com/outlook/add-ins/quick-start)