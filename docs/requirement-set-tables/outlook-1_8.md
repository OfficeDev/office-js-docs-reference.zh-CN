| 类 | 域 | 说明 |
|:---|:---|:---|
|[AppointmentCompose](/javascript/api/outlook/outlook.appointmentcompose)|[addFileAttachmentFromBase64Async (base64File： string， attachmentName： string， options？： Office.AsyncContextOptions & { isInline： boolean }， callback？： (asyncResult： Office.AsyncResult <string>) => void) ](/javascript/api/outlook/outlook.appointmentcompose#addfileattachmentfrombase64async-base64file--attachmentname--options--isinline--callback--asyncresult-)|将文件作为附件添加到邮件或约会。|
||[categories](/javascript/api/outlook/outlook.appointmentcompose#categories)|获取一个对象，该对象提供用于管理项目类别的方法。|
||[enhancedLocation](/javascript/api/outlook/outlook.appointmentcompose#enhancedlocation)|获取或设置约会的位置。|
||[getAttachmentContentAsync (attachmentId： string， options？： Office.AsyncContextOptions， callback？： (asyncResult： Office.AsyncResult <AttachmentContent>) => void) ](/javascript/api/outlook/outlook.appointmentcompose#getattachmentcontentasync-attachmentid--options--callback--asyncresult-)|从邮件或约会获取附件，并作为对象 `AttachmentContent` 返回。|
||[getAttachmentsAsync (选项？： Office.AsyncContextOptions， callback？： (asyncResult： Office.AsyncResult<AttachmentDetailsCompose[]>) => void) ](/javascript/api/outlook/outlook.appointmentcompose#getattachmentsasync-options--callback--asyncresult-)|获取项目的附件作为数组。|
||[getItemIdAsync (回调： (asyncResult： Office.AsyncResult) <string> => void) ](/javascript/api/outlook/outlook.appointmentcompose#getitemidasync-callback--asyncresult-)|异步获取已保存项目的 ID。|
||[getItemIdAsync (选项：Office.AsyncContextOptions，回调： (asyncResult： Office.AsyncResult) <string> => void) ](/javascript/api/outlook/outlook.appointmentcompose#getitemidasync-options--callback--asyncresult-)|异步获取已保存项目的 ID。|
||[getSharedPropertiesAsync (回调： (asyncResult： Office.AsyncResult <SharedProperties>) => void) ](/javascript/api/outlook/outlook.appointmentcompose#getsharedpropertiesasync-callback--asyncresult-)|获取共享文件夹中约会或邮件的属性。|
||[getSharedPropertiesAsync (选项：Office.AsyncContextOptions，回调： (asyncResult： Office.AsyncResult <SharedProperties>) => void) ](/javascript/api/outlook/outlook.appointmentcompose#getsharedpropertiesasync-options--callback--asyncresult-)|获取共享文件夹中约会或邮件的属性。|
|[AppointmentRead](/javascript/api/outlook/outlook.appointmentread)|[categories](/javascript/api/outlook/outlook.appointmentread#categories)|获取一个对象，该对象提供用于管理项目类别的方法。|
||[enhancedLocation](/javascript/api/outlook/outlook.appointmentread#enhancedlocation)|获取约会的位置。|
||[getAttachmentContentAsync (attachmentId： string， options？： Office.AsyncContextOptions， callback？： (asyncResult： Office.AsyncResult <AttachmentContent>) => void) ](/javascript/api/outlook/outlook.appointmentread#getattachmentcontentasync-attachmentid--options--callback--asyncresult-)|从邮件或约会获取附件，并作为对象 `AttachmentContent` 返回。|
||[getSharedPropertiesAsync (回调： (asyncResult： Office.AsyncResult <SharedProperties>) => void) ](/javascript/api/outlook/outlook.appointmentread#getsharedpropertiesasync-callback--asyncresult-)|获取共享文件夹中约会或邮件的属性。|
||[getSharedPropertiesAsync (选项：Office.AsyncContextOptions，回调： (asyncResult： Office.AsyncResult <SharedProperties>) => void) ](/javascript/api/outlook/outlook.appointmentread#getsharedpropertiesasync-options--callback--asyncresult-)|获取共享文件夹中约会或邮件的属性。|
|[AttachmentContent](/javascript/api/outlook/outlook.attachmentcontent)|[content](/javascript/api/outlook/outlook.attachmentcontent#content)|作为字符串的附件内容。|
||[format](/javascript/api/outlook/outlook.attachmentcontent#format)|用于附件内容的字符串格式。|
|[AttachmentDetailsCompose](/javascript/api/outlook/outlook.attachmentdetailscompose)|[attachmentType](/javascript/api/outlook/outlook.attachmentdetailscompose#attachmenttype)|获取一个指示附件类型的值。|
||[id](/javascript/api/outlook/outlook.attachmentdetailscompose#id)|获取附件的索引。|
||[isInline](/javascript/api/outlook/outlook.attachmentdetailscompose#isinline)|获取指示是否应在项目正文中显示附件的值。|
||[名称](/javascript/api/outlook/outlook.attachmentdetailscompose#name)|获取附件的名称。|
||[size](/javascript/api/outlook/outlook.attachmentdetailscompose#size)|获取以字节为单位的附件大小。|
||[url](/javascript/api/outlook/outlook.attachmentdetailscompose#url)|获取附件的 URL（如果其类型为 `MailboxEnums.AttachmentType.Cloud` 。|
|[AttachmentsChangedEventArgs](/javascript/api/outlook/outlook.attachmentschangedeventargs)|[attachmentDetails](/javascript/api/outlook/outlook.attachmentschangedeventargs#attachmentdetails)||
||[attachmentStatus](/javascript/api/outlook/outlook.attachmentschangedeventargs#attachmentstatus)|获取附件是已添加还是已删除。|
||[type](/javascript/api/outlook/outlook.attachmentschangedeventargs#type)|获取事件的类型。|
|[Categories](/javascript/api/outlook/outlook.categories)|[addAsync (类别： string[]， options？： Office.AsyncContextOptions， callback？： (asyncResult： Office.AsyncResult <void>) => void) ](/javascript/api/outlook/outlook.categories#addasync-categories--options--callback--asyncresult-)|向项目添加类别。|
||[getAsync (： (asyncResult： Office.AsyncResult<CategoryDetails[]>) => void) ](/javascript/api/outlook/outlook.categories#getasync-callback--asyncresult-)|获取项目的类别。|
||[getAsync (选项： Office.AsyncContextOptions， callback： (asyncResult： Office.AsyncResult<CategoryDetails[]>) => void) ](/javascript/api/outlook/outlook.categories#getasync-options--callback--asyncresult-)|获取项目的类别。|
||[removeAsync (categories： string[]， options？： Office.AsyncContextOptions， callback？： (asyncResult： Office.AsyncResult <void>) => void) ](/javascript/api/outlook/outlook.categories#removeasync-categories--options--callback--asyncresult-)|从项目中删除类别。|
|[CategoryDetails](/javascript/api/outlook/outlook.categorydetails)|[color](/javascript/api/outlook/outlook.categorydetails#color)|类别的颜色。|
||[displayName](/javascript/api/outlook/outlook.categorydetails#displayname)|类别的名称。|
|[EnhancedLocation](/javascript/api/outlook/outlook.enhancedlocation)|[addAsync (locationIdentifiers： LocationIdentifier[]， options？： Office.AsyncContextOptions， callback？： (asyncResult： Office.AsyncResult <void>) => void) ](/javascript/api/outlook/outlook.enhancedlocation#addasync-locationidentifiers--options--callback--asyncresult-)|添加到与约会关联的位置集。|
||[getAsync (选项？： Office.AsyncContextOptions， callback？： (asyncResult： Office.AsyncResult<LocationDetails[]>) => void) ](/javascript/api/outlook/outlook.enhancedlocation#getasync-options--callback--asyncresult-)|获取与约会关联的位置集。|
||[removeAsync (locationIdentifiers： LocationIdentifier[]， options？： Office.AsyncContextOptions， callback？： (asyncResult： Office.AsyncResult <void>) => void) ](/javascript/api/outlook/outlook.enhancedlocation#removeasync-locationidentifiers--options--callback--asyncresult-)|删除与约会关联的位置集。|
|[EnhancedLocationsChangedEventArgs](/javascript/api/outlook/outlook.enhancedlocationschangedeventargs)|[enhancedLocations](/javascript/api/outlook/outlook.enhancedlocationschangedeventargs#enhancedlocations)|获取增强位置集。|
||[type](/javascript/api/outlook/outlook.enhancedlocationschangedeventargs#type)|获取事件的类型。|
|[InternetHeaders](/javascript/api/outlook/outlook.internetheaders)|[getAsync (： string[]， options？： Office.AsyncContextOptions， callback？： (asyncResult： Office.AsyncResult <InternetHeaders>) => void) ](/javascript/api/outlook/outlook.internetheaders#getasync-names--options--callback--asyncresult-)|如果给定一组 Internet 标头名称，此方法将返回包含这些 Internet 标头及其值的字典。|
||[removeAsync (： string[]， options？： Office.AsyncContextOptions， callback？： (asyncResult： Office.AsyncResult <InternetHeaders>) => void) ](/javascript/api/outlook/outlook.internetheaders#removeasync-names--options--callback--asyncresult-)|如果给定一组 Internet 标头名称，此方法将删除 Internet 标头集合中的指定标头。|
||[setAsync (标头： 对象， 选项？： Office.AsyncContextOptions， callback？： (asyncResult： Office.AsyncResult <void>) => void) ](/javascript/api/outlook/outlook.internetheaders#setasync-headers--options--callback--asyncresult-)|将指定的 Internet 标头设置为指定的值。|
|[LocationDetails](/javascript/api/outlook/outlook.locationdetails)|[displayName](/javascript/api/outlook/outlook.locationdetails#displayname)|位置显示名称。|
||[emailAddress](/javascript/api/outlook/outlook.locationdetails#emailaddress)|与位置关联的电子邮件地址。|
||[locationIdentifier](/javascript/api/outlook/outlook.locationdetails#locationidentifier)|`LocationIdentifier`位置的位置。|
|[LocationIdentifier](/javascript/api/outlook/outlook.locationidentifier)|[id](/javascript/api/outlook/outlook.locationidentifier#id)|位置的唯一 ID。|
||[type](/javascript/api/outlook/outlook.locationidentifier#type)|位置的类型。|
|[邮箱](/javascript/api/outlook/outlook.mailbox)|[masterCategories](/javascript/api/outlook/outlook.mailbox#mastercategories)|获取一个对象，该对象提供用于管理与邮箱关联的类别主列表的方法。|
|[MasterCategories](/javascript/api/outlook/outlook.mastercategories)|[addAsync (类别： CategoryDetails[]， options？： Office.AsyncContextOptions， callback？： (asyncResult： Office.AsyncResult <void>) => void) ](/javascript/api/outlook/outlook.mastercategories#addasync-categories--options--callback--asyncresult-)|将类别添加到邮箱的主列表。|
||[getAsync (： (asyncResult： Office.AsyncResult<CategoryDetails[]>) => void) ](/javascript/api/outlook/outlook.mastercategories#getasync-callback--asyncresult-)|获取邮箱上的类别主列表。|
||[getAsync (选项： Office.AsyncContextOptions， callback： (asyncResult： Office.AsyncResult<CategoryDetails[]>) => void) ](/javascript/api/outlook/outlook.mastercategories#getasync-options--callback--asyncresult-)|获取邮箱上的类别主列表。|
||[removeAsync (categories： string[]， options？： Office.AsyncContextOptions， callback？： (asyncResult： Office.AsyncResult <void>) => void) ](/javascript/api/outlook/outlook.mastercategories#removeasync-categories--options--callback--asyncresult-)|从邮箱的主列表中删除类别。|
|[MessageCompose](/javascript/api/outlook/outlook.messagecompose)|[addFileAttachmentFromBase64Async (base64File： string， attachmentName： string， options？： Office.AsyncContextOptions & { isInline： boolean }， callback？： (asyncResult： Office.AsyncResult <string>) => void) ](/javascript/api/outlook/outlook.messagecompose#addfileattachmentfrombase64async-base64file--attachmentname--options--isinline--callback--asyncresult-)|将文件作为附件添加到邮件或约会。|
||[categories](/javascript/api/outlook/outlook.messagecompose#categories)|获取一个对象，该对象提供用于管理项目类别的方法。|
||[getAttachmentContentAsync (attachmentId： string， options？： Office.AsyncContextOptions， callback？： (asyncResult： Office.AsyncResult <AttachmentContent>) => void) ](/javascript/api/outlook/outlook.messagecompose#getattachmentcontentasync-attachmentid--options--callback--asyncresult-)|从邮件或约会获取附件，并作为对象 `AttachmentContent` 返回。|
||[getAttachmentsAsync (选项？： Office.AsyncContextOptions， callback？： (asyncResult： Office.AsyncResult<AttachmentDetailsCompose[]>) => void) ](/javascript/api/outlook/outlook.messagecompose#getattachmentsasync-options--callback--asyncresult-)|获取项目的附件作为数组。|
||[getItemIdAsync (回调： (asyncResult： Office.AsyncResult) <string> => void) ](/javascript/api/outlook/outlook.messagecompose#getitemidasync-callback--asyncresult-)|异步获取已保存项目的 ID。|
||[getItemIdAsync (选项：Office.AsyncContextOptions，回调： (asyncResult： Office.AsyncResult) <string> => void) ](/javascript/api/outlook/outlook.messagecompose#getitemidasync-options--callback--asyncresult-)|异步获取已保存项目的 ID。|
||[getSharedPropertiesAsync (回调： (asyncResult： Office.AsyncResult <SharedProperties>) => void) ](/javascript/api/outlook/outlook.messagecompose#getsharedpropertiesasync-callback--asyncresult-)|获取共享文件夹中约会或邮件的属性。|
||[getSharedPropertiesAsync (选项：Office.AsyncContextOptions，回调： (asyncResult： Office.AsyncResult <SharedProperties>) => void) ](/javascript/api/outlook/outlook.messagecompose#getsharedpropertiesasync-options--callback--asyncresult-)|获取共享文件夹中约会或邮件的属性。|
||[internetHeaders](/javascript/api/outlook/outlook.messagecompose#internetheaders)|获取或设置邮件的自定义 Internet 标头。|
|[MessageRead](/javascript/api/outlook/outlook.messageread)|[categories](/javascript/api/outlook/outlook.messageread#categories)|获取一个对象，该对象提供用于管理项目类别的方法。|
||[getAllInternetHeadersAsync (选项？： Office.AsyncContextOptions， callback？： (asyncResult： Office.AsyncResult <string>) => void) ](/javascript/api/outlook/outlook.messageread#getallinternetheadersasync-options--callback--asyncresult-)|以字符串形式获取邮件的所有 Internet 标头。|
||[getAttachmentContentAsync (attachmentId： string， options？： Office.AsyncContextOptions， callback？： (asyncResult： Office.AsyncResult <AttachmentContent>) => void) ](/javascript/api/outlook/outlook.messageread#getattachmentcontentasync-attachmentid--options--callback--asyncresult-)|从邮件或约会获取附件，并作为对象 `AttachmentContent` 返回。|
||[getSharedPropertiesAsync (回调： (asyncResult： Office.AsyncResult <SharedProperties>) => void) ](/javascript/api/outlook/outlook.messageread#getsharedpropertiesasync-callback--asyncresult-)|获取共享文件夹中约会或邮件的属性。|
||[getSharedPropertiesAsync (选项：Office.AsyncContextOptions，回调： (asyncResult： Office.AsyncResult <SharedProperties>) => void) ](/javascript/api/outlook/outlook.messageread#getsharedpropertiesasync-options--callback--asyncresult-)|获取共享文件夹中约会或邮件的属性。|
|[SharedProperties](/javascript/api/outlook/outlook.sharedproperties)|[delegatePermissions](/javascript/api/outlook/outlook.sharedproperties#delegatepermissions)|代理对共享文件夹拥有的权限。|
||[owner](/javascript/api/outlook/outlook.sharedproperties#owner)|共享项目的所有者的电子邮件地址。|
||[targetMailbox](/javascript/api/outlook/outlook.sharedproperties#targetmailbox)|委派访问的所有者邮箱的位置。|
||[targetRestUrl](/javascript/api/outlook/outlook.sharedproperties#targetresturl)|REST API 的基本 URL (当前 https://outlook.office.com/api) 。|
