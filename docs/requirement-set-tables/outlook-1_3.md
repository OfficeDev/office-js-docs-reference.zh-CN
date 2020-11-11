| Class | 域 | 说明 |
|:---|:---|:---|
|[AppointmentCompose](/javascript/api/outlook/outlook.appointmentcompose)|[关闭 ( # B1 ](/javascript/api/outlook/outlook.appointmentcompose#close--)|关闭正在撰写的当前项目|
||[notificationMessages](/javascript/api/outlook/outlook.appointmentcompose#notificationmessages)|获取项目的通知邮件。|
||[saveAsync (回呼： (asyncResult： <string>) => void) ](/javascript/api/outlook/outlook.appointmentcompose#saveasync-callback--asyncresult-)|异步保存项目。|
||[saveAsync (选项： AsyncContextOptions、回拨： (asyncResult： <string>) => void) ](/javascript/api/outlook/outlook.appointmentcompose#saveasync-options--callback--asyncresult-)|异步保存项目。|
|[AppointmentRead](/javascript/api/outlook/outlook.appointmentread)|[notificationMessages](/javascript/api/outlook/outlook.appointmentread#notificationmessages)|获取项目的通知邮件。|
|[正文](/javascript/api/outlook/outlook.body)|[getAsync (coercionType： CoercionType \| string，options？： AsyncContextOptions，callback？： (asyncResult： <string>) => void) ](/javascript/api/outlook/outlook.body#getasync-coerciontype--options--callback--asyncresult-)|使用指定的格式返回当前正文。|
||[setAsync (数据： string，options？： AsyncContextOptions & CoercionTypeOptions，callback？： (asyncResult： <void>) => void) ](/javascript/api/outlook/outlook.body#setasync-data--options--callback--asyncresult-)|将整个正文替换为指定的文本。|
|[邮箱](/javascript/api/outlook/outlook.mailbox)|[Office.context.mailbox.converttoewsid (itemId： string，Office.mailboxenums.restversion： MailboxEnums \| string) ](/javascript/api/outlook/outlook.mailbox#converttoewsid-itemid--restversion-)|将项目 ID 格式化（从 REST 转换为 EWS 格式）。|
||[Office.context.mailbox.converttorestid (itemId： string，Office.mailboxenums.restversion： MailboxEnums \| string) ](/javascript/api/outlook/outlook.mailbox#converttorestid-itemid--restversion-)|将项目 ID 格式化（从 EWS 转换为 REST 格式）。|
|[MessageCompose](/javascript/api/outlook/outlook.messagecompose)|[关闭 ( # B1 ](/javascript/api/outlook/outlook.messagecompose#close--)|关闭正在撰写的当前项目|
||[notificationMessages](/javascript/api/outlook/outlook.messagecompose#notificationmessages)|获取项目的通知邮件。|
||[saveAsync (回呼： (asyncResult： <string>) => void) ](/javascript/api/outlook/outlook.messagecompose#saveasync-callback--asyncresult-)|异步保存项目。|
||[saveAsync (选项： AsyncContextOptions、回拨： (asyncResult： <string>) => void) ](/javascript/api/outlook/outlook.messagecompose#saveasync-options--callback--asyncresult-)|异步保存项目。|
|[MessageRead](/javascript/api/outlook/outlook.messageread)|[notificationMessages](/javascript/api/outlook/outlook.messageread#notificationmessages)|获取项目的通知邮件。|
|[NotificationMessageDetails](/javascript/api/outlook/outlook.notificationmessagedetails)|[icon](/javascript/api/outlook/outlook.notificationmessagedetails#icon)|对在清单的 `Resources` 部分中定义的图标的引用。|
||[key](/javascript/api/outlook/outlook.notificationmessagedetails#key)|通知邮件标识符。|
||[message](/javascript/api/outlook/outlook.notificationmessagedetails#message)|通知邮件的文本。|
||[保持](/javascript/api/outlook/outlook.notificationmessagedetails#persistent)|指定邮件是否应为永久性。|
||[type](/javascript/api/outlook/outlook.notificationmessagedetails#type)|指定 `ItemNotificationMessageType` 邮件。|
|[NotificationMessages](/javascript/api/outlook/outlook.notificationmessages)|[addAsync (key： string，JSONmessage： NotificationMessageDetails，options？： AsyncContextOptions，callback？： (asyncResult： <void>) => void) ](/javascript/api/outlook/outlook.notificationmessages#addasync-key--jsonmessage--options--callback--asyncresult-)|向项目添加通知。|
||[getAllAsync (选项？： AsyncContextOptions、callback？： (asyncResult： NotificationMessageDetails<[] >) => void) ](/javascript/api/outlook/outlook.notificationmessages#getallasync-options--callback--asyncresult-)|返回某个项目的所有项和邮件。|
||[removeAsync (key： string，options？： AsyncContextOptions，回呼？： (asyncResult： <void>) => void) ](/javascript/api/outlook/outlook.notificationmessages#removeasync-key--options--callback--asyncresult-)|删除某个项目的通知邮件。|
||[replaceAsync (key： string，JSONmessage： NotificationMessageDetails，options？： AsyncContextOptions，callback？： (asyncResult： <void>) => void) ](/javascript/api/outlook/outlook.notificationmessages#replaceasync-key--jsonmessage--options--callback--asyncresult-)|将带有给定项的通知邮件替换为另一封邮件。|
