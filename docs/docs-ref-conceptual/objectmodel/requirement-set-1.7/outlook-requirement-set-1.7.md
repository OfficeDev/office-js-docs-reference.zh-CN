# <a name="outlook-add-in-api-requirement-set-17"></a>Outlook 加载项 API 要求集 1.7

适用于 Office 的 JavaScript API 的 Outlook 外接程序 API 子集包括您可以在 Outlook 外接程序中使用的对象、方法、属性和事件。

## <a name="whats-new-in-17"></a>What's new in 1.7？

要求集 1.7 包括所有[要求](../requirement-set-1.6/outlook-requirement-set-1.6.md)集 1.6 的功能。 它添加以下功能。

- 添加了有关定期模式的约会和会议请求的消息的新 Api。
- 修改 item.from 属性，也是在撰写模式下可用。
- 添加了的支持 RecurrenceChanged、 RecipientsChanged 和 AppointmentTimeChanged 事件。

### <a name="change-log"></a>更改日志

- [从](/javascript/api/outlook_1_7/office.from)添加： 添加新的对象，提供一个方法来获取值。
- 添加[组织者](/javascript/api/outlook_1_7/office.organizer)： 添加新的对象，提供要获取组织者值的方法。
- 添加[定期](/javascript/api/outlook_1_7/office.recurrence)： 添加新的对象，提供方法获取和设置约会的定期模式，但仅的获取消息的定期模式会议请求。
- 添加[RecurrenceTimeZone](/javascript/api/outlook_1_7/office.recurrencetimezone)： 添加新的对象值，该值代表定期模式的时区配置。
- 添加[SeriesTime](/javascript/api/outlook_1_7/office.seriestime)： 添加新的对象提供获取和设置定期系列中的日期和时间的约会，并可获取的日期和时间的会议请求定期系列中的方法。
- 添加[Office.context.mailbox.item.addHandlerAsync](office.context.mailbox.item.md#addhandlerasynceventtype-handler-options-callback)： 添加用于添加事件处理程序支持的事件的新方法。
- 修改[Office.context.mailbox.item.from](office.context.mailbox.item.md#from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom)： 修改要获取在撰写模式下的值。
- 修改的[Office.context.mailbox.item.organizer](office.context.mailbox.item.md#organizer-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsorganizerjavascriptapioutlook17officeorganizer) -修改在撰写模式下获取组织者值。
- 添加[Office.context.mailbox.item.recurrence](office.context.mailbox.item.md#nullable-recurrence-recurrencejavascriptapioutlook17officerecurrence)： 添加新属性来获取或设置一个值，它提供用于管理定期模式的约会项目的方法。 此属性还可用于获取定期模式的会议请求项目。
- 添加[Office.context.mailbox.item.removeHandlerAsync](office.context.mailbox.item.md#removehandlerasynceventtype-handler-options-callback)： 添加事件处理程序中删除的新方法。
- 添加[Office.context.mailbox.item.seriesId](office.context.mailbox.item.md#nullable-seriesid-string)： 添加所属获取匹配项的数据系列的 id 的新属性。
- 添加[Office.MailboxEnums.Days](/javascript/api/outlook_1_7/office.mailboxenums.days)： 添加指定的天的周或一天中的类型的新枚举。
- 添加[Office.MailboxEnums.Month](/javascript/api/outlook_1_7/office.mailboxenums.month)： 添加指定的月份新枚举。
- 添加[Office.MailboxEnums.RecurrenceTimeZone](/javascript/api/outlook_1_7/office.mailboxenums.recurrencetimezone)： 添加指定时区定期应用于新枚举。
- 添加[Office.MailboxEnums.RecurrenceType](/javascript/api/outlook_1_7/office.mailboxenums.recurrencetype)： 添加新的枚举，指定重复的类型。
- 添加[Office.MailboxEnums.WeekNumber](/javascript/api/outlook_1_7/office.mailboxenums.weeknumber)： 添加指定月的周的新枚举。
- 修改[Office.EventType](/javascript/api/office/office.eventtype)： 修改要支持通过添加 RecurrenceChanged、 RecipientsChanged 和 AppointmentTimeChanged 事件`RecurrenceChanged`， `RecipientsChanged`，和`AppointmentTimeChanged`条目分别。

## <a name="see-also"></a>另请参阅

- [Outlook 外接程序](https://docs.microsoft.com/outlook/add-ins/)
- [Outlook 外接程序代码示例](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [入门](https://docs.microsoft.com/outlook/add-ins/quick-start)