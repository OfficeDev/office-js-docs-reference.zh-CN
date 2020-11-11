| Class | 域 | 说明 |
|:---|:---|:---|
|[AppointmentCompose](/javascript/api/outlook/outlook.appointmentcompose)|[addFileAttachmentAsync (uri： string，attachmentName： string，options？： AsyncContextOptions & {isInline： boolean}，callback？： (asyncResult： <string>) => void) ](/javascript/api/outlook/outlook.appointmentcompose#addfileattachmentasync-uri--attachmentname--options--isinline--callback--asyncresult-)|将文件作为附件添加到邮件或约会。|
||[addItemAttachmentAsync (itemId： any，attachmentName： string，options？： AsyncContextOptions，callback？： (asyncResult： <string>) => void) ](/javascript/api/outlook/outlook.appointmentcompose#additemattachmentasync-itemid--attachmentname--options--callback--asyncresult-)|将 Exchange 项目（如邮件）作为附件添加到邮件或约会。|
||[body](/javascript/api/outlook/outlook.appointmentcompose#body)|获取一个提供用于处理项目正文的方法的对象。|
||[isInline](/javascript/api/outlook/outlook.appointmentcompose#isinline)||
||[removeAttachmentAsync (attachmentId： string，options？： AsyncContextOptions，callback？： (asyncResult： <void>) => void) ](/javascript/api/outlook/outlook.appointmentcompose#removeattachmentasync-attachmentid--options--callback--asyncresult-)|将附件从邮件或约会中删除。|
|[AppointmentForm](/javascript/api/outlook/outlook.appointmentform)|[body](/javascript/api/outlook/outlook.appointmentform#body)|获取一个提供用于处理项目正文的方法的对象。|
|[AppointmentRead](/javascript/api/outlook/outlook.appointmentread)|[body](/javascript/api/outlook/outlook.appointmentread#body)|获取一个提供用于处理项目正文的方法的对象。|
|[AttachmentDetails](/javascript/api/outlook/outlook.attachmentdetails)|[attachmentType](/javascript/api/outlook/outlook.attachmentdetails#attachmenttype)|获取一个指示附件类型的值。|
||[contentType](/javascript/api/outlook/outlook.attachmentdetails#contenttype)|获取附件的 MIME 内容类型。|
||[id](/javascript/api/outlook/outlook.attachmentdetails#id)|获取附件的 Exchange 附件 ID。|
||[isInline](/javascript/api/outlook/outlook.attachmentdetails#isinline)|获取指示是否应在项目正文中显示附件的值。|
||[name](/javascript/api/outlook/outlook.attachmentdetails#name)|获取附件的名称。|
||[size](/javascript/api/outlook/outlook.attachmentdetails#size)|获取以字节为单位的附件大小。|
|[正文](/javascript/api/outlook/outlook.body)|[getTypeAsync (选项？： AsyncContextOptions、callback？： (asyncResult：<CoercionType>) => void) ](/javascript/api/outlook/outlook.body#gettypeasync-options--callback--asyncresult-)|获取一个值，该值指示内容采用 HTML 格式还是文本格式。|
||[prependAsync (数据： string，options？： AsyncContextOptions & CoercionTypeOptions，callback？： (asyncResult： <void>) => void) ](/javascript/api/outlook/outlook.body#prependasync-data--options--callback--asyncresult-)|将指定内容添加到项目正文开头。|
||[Document.setselecteddataasync (数据： string，options？： AsyncContextOptions & CoercionTypeOptions，callback？： (asyncResult： <void>) => void) ](/javascript/api/outlook/outlook.body#setselecteddataasync-data--options--callback--asyncresult-)|将正文中的所选内容更换为指定文本。|
|[位置](/javascript/api/outlook/outlook.location)|[getAsync (回呼： (asyncResult： <string>) => void) ](/javascript/api/outlook/outlook.location#getasync-callback--asyncresult-)|获取约会的位置。|
||[getAsync (选项： AsyncContextOptions、回拨： (asyncResult： <string>) => void) ](/javascript/api/outlook/outlook.location#getasync-options--callback--asyncresult-)|获取约会的位置。|
||[setAsync (location： string，options？： AsyncContextOptions，回呼？： (asyncResult： <void>) => void) ](/javascript/api/outlook/outlook.location#setasync-location--options--callback--asyncresult-)|设置约会的位置。|
|[MessageCompose](/javascript/api/outlook/outlook.messagecompose)|[addFileAttachmentAsync (uri： string，attachmentName： string，options？： AsyncContextOptions & {isInline： boolean}，callback？： (asyncResult： <string>) => void) ](/javascript/api/outlook/outlook.messagecompose#addfileattachmentasync-uri--attachmentname--options--isinline--callback--asyncresult-)|将文件作为附件添加到邮件或约会。|
||[addItemAttachmentAsync (itemId： any，attachmentName： string，options？： AsyncContextOptions，callback？： (asyncResult： <string>) => void) ](/javascript/api/outlook/outlook.messagecompose#additemattachmentasync-itemid--attachmentname--options--callback--asyncresult-)|将 Exchange 项目（如邮件）作为附件添加到邮件或约会。|
||[bcc](/javascript/api/outlook/outlook.messagecompose#bcc)|获取一个对象，该对象提供用于获取或更新邮件的 **密件抄送** (密件抄送) 行上的收件人的方法。|
||[body](/javascript/api/outlook/outlook.messagecompose#body)|获取一个提供用于处理项目正文的方法的对象。|
||[isInline](/javascript/api/outlook/outlook.messagecompose#isinline)||
||[removeAttachmentAsync (attachmentId： string，options？： AsyncContextOptions，callback？： (asyncResult： <void>) => void) ](/javascript/api/outlook/outlook.messagecompose#removeattachmentasync-attachmentid--options--callback--asyncresult-)|将附件从邮件或约会中删除。|
|[MessageRead](/javascript/api/outlook/outlook.messageread)|[body](/javascript/api/outlook/outlook.messageread#body)|获取一个提供用于处理项目正文的方法的对象。|
|[PhoneNumber](/javascript/api/outlook/outlook.phonenumber)|[type](/javascript/api/outlook/outlook.phonenumber#type)|获取标识电话号码的类型的字符串： Home、Work、Mobile、未指定。|
|[收件人](/javascript/api/outlook/outlook.recipients)|[addAsync (收件人： (string \| Emailuser.displayname \| EmailAddressDetails) []，options？： AsyncContextOptions，回呼？： (asyncResult： <void>) => void) ](/javascript/api/outlook/outlook.recipients#addasync-recipients-)|将收件人列表添加到约会或邮件的现有收件人中。|
||[getAsync (回呼： (asyncResult： EmailAddressDetails<[] >) => void) ](/javascript/api/outlook/outlook.recipients#getasync-callback--asyncresult-)|获取约会或邮件的收件人列表。|
||[getAsync (选项： AsyncContextOptions、回拨： (asyncResult：<EmailAddressDetails [] >) => void) ](/javascript/api/outlook/outlook.recipients#getasync-options--callback--asyncresult-)|获取约会或邮件的收件人列表。|
||[setAsync (收件人： (string \| Emailuser.displayname \| EmailAddressDetails) []，callback： (asyncResult： <void>) => void) ](/javascript/api/outlook/outlook.recipients#setasync-recipients-)|设置约会或邮件的收件人列表。|
||[setAsync (收件人： (string \| Emailuser.displayname \| EmailAddressDetails) []，Options： AsyncContextOptions，callback： (asyncResult： <void>) => void) ](/javascript/api/outlook/outlook.recipients#setasync-recipients-)|设置约会或邮件的收件人列表。|
|[主题](/javascript/api/outlook/outlook.subject)|[getAsync (回呼： (asyncResult： <string>) => void) ](/javascript/api/outlook/outlook.subject#getasync-callback--asyncresult-)|获取约会或邮件的主题。|
||[getAsync (选项： AsyncContextOptions、回拨： (asyncResult： <string>) => void) ](/javascript/api/outlook/outlook.subject#getasync-options--callback--asyncresult-)|获取约会或邮件的主题。|
||[setAsync (subject： string，options？： AsyncContextOptions，callback？： (asyncResult： <void>) => void) ](/javascript/api/outlook/outlook.subject#setasync-subject--options--callback--asyncresult-)|设置约会或邮件的主题。|
|[Time](/javascript/api/outlook/outlook.time)|[getAsync (回呼： (asyncResult： <Date>) => void) ](/javascript/api/outlook/outlook.time#getasync-callback--asyncresult-)|获取约会的开始或结束时间。|
||[getAsync (选项： AsyncContextOptions、回拨： (asyncResult： <Date>) => void) ](/javascript/api/outlook/outlook.time#getasync-options--callback--asyncresult-)|获取约会的开始或结束时间。|
||[setAsync (dateTime： Date、options？： AsyncContextOptions、callback？： (asyncResult： <void>) => void) ](/javascript/api/outlook/outlook.time#setasync-datetime--options--callback--asyncresult-)|设置约会的开始或结束时间。|
