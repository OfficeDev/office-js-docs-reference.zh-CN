| Class | 域 | 说明 |
|:---|:---|:---|
|[AppointmentCompose](/javascript/api/outlook/outlook.appointmentcompose)|[addHandlerAsync (事件类型： Office. 事件 \| 类型字符串、处理程序： any、options？： AsyncContextOptions、callback？： (asyncResult： <void>) => void) ](/javascript/api/outlook/outlook.appointmentcompose#addhandlerasync-eventtype--handler--options--callback--asyncresult-)|添加支持事件的事件处理程序。|
||[organizer](/javascript/api/outlook/outlook.appointmentcompose#organizer)|获取指定会议的组织者。|
||[循环](/javascript/api/outlook/outlook.appointmentcompose#recurrence)|获取或设置约会的定期模式。|
||[removeHandlerAsync (事件类型： Office. 命令 \| 行字符串、选项？： AsyncContextOptions、callback？： (asyncResult： <void>) => void) ](/javascript/api/outlook/outlook.appointmentcompose#removehandlerasync-eventtype--options--callback--asyncresult-)|删除受支持事件类型的事件处理程序。|
||[Webcasts&seriesid](/javascript/api/outlook/outlook.appointmentcompose#seriesid)|获取实例所属的系列的 id。|
|[AppointmentRead](/javascript/api/outlook/outlook.appointmentread)|[addHandlerAsync (事件类型： Office. 事件 \| 类型字符串、处理程序： any、options？： AsyncContextOptions、callback？： (asyncResult： <void>) => void) ](/javascript/api/outlook/outlook.appointmentread#addhandlerasync-eventtype--handler--options--callback--asyncresult-)|添加支持事件的事件处理程序。|
||[循环](/javascript/api/outlook/outlook.appointmentread#recurrence)|获取约会的定期模式。|
||[removeHandlerAsync (事件类型： Office. 命令 \| 行字符串、选项？： AsyncContextOptions、callback？： (asyncResult： <void>) => void) ](/javascript/api/outlook/outlook.appointmentread#removehandlerasync-eventtype--options--callback--asyncresult-)|删除受支持事件类型的事件处理程序。|
||[Webcasts&seriesid](/javascript/api/outlook/outlook.appointmentread#seriesid)|获取实例所属的系列的 ID。|
|[AppointmentTimeChangedEventArgs](/javascript/api/outlook/outlook.appointmenttimechangedeventargs)|[end](/javascript/api/outlook/outlook.appointmenttimechangedeventargs#end)||
||[start](/javascript/api/outlook/outlook.appointmenttimechangedeventargs#start)||
||[type](/javascript/api/outlook/outlook.appointmenttimechangedeventargs#type)||
|[From](/javascript/api/outlook/outlook.from)|[getAsync (选项？： AsyncContextOptions，回呼？： (asyncResult： <EmailAddressDetails>) => void) ](/javascript/api/outlook/outlook.from#getasync-options--callback--asyncresult-)|获取邮件的起始值。|
|[MessageCompose](/javascript/api/outlook/outlook.messagecompose)|[addHandlerAsync (事件类型： Office. 事件 \| 类型字符串、处理程序： any、options？： AsyncContextOptions、callback？： (asyncResult： <void>) => void) ](/javascript/api/outlook/outlook.messagecompose#addhandlerasync-eventtype--handler--options--callback--asyncresult-)|添加支持事件的事件处理程序。|
||[from](/javascript/api/outlook/outlook.messagecompose#from)|获取邮件发件人的电子邮件地址。|
||[removeHandlerAsync (事件类型： Office. 命令 \| 行字符串、选项？： AsyncContextOptions、callback？： (asyncResult： <void>) => void) ](/javascript/api/outlook/outlook.messagecompose#removehandlerasync-eventtype--options--callback--asyncresult-)|删除受支持事件类型的事件处理程序。|
||[Webcasts&seriesid](/javascript/api/outlook/outlook.messagecompose#seriesid)|获取实例所属的系列的 ID。|
|[MessageRead](/javascript/api/outlook/outlook.messageread)|[addHandlerAsync (事件类型： Office. 事件 \| 类型字符串、处理程序： any、options？： AsyncContextOptions、callback？： (asyncResult： <void>) => void) ](/javascript/api/outlook/outlook.messageread#addhandlerasync-eventtype--handler--options--callback--asyncresult-)|添加支持事件的事件处理程序。|
||[循环](/javascript/api/outlook/outlook.messageread#recurrence)|获取约会的定期模式。|
||[removeHandlerAsync (事件类型： Office. 命令 \| 行字符串、选项？： AsyncContextOptions、callback？： (asyncResult： <void>) => void) ](/javascript/api/outlook/outlook.messageread#removehandlerasync-eventtype--options--callback--asyncresult-)|删除受支持事件类型的事件处理程序。|
||[Webcasts&seriesid](/javascript/api/outlook/outlook.messageread#seriesid)|获取实例所属的系列的 id。|
|[Organizer](/javascript/api/outlook/outlook.organizer)|[getAsync (选项？： AsyncContextOptions，回呼？： (asyncResult： <EmailAddressDetails>) => void) ](/javascript/api/outlook/outlook.organizer#getasync-options--callback--asyncresult-)|将约会的组织者值获取为 {@link EmailAddressDetails | EmailAddressDetails 对象|
|[RecipientsChangedEventArgs](/javascript/api/outlook/outlook.recipientschangedeventargs)|[changedRecipientFields](/javascript/api/outlook/outlook.recipientschangedeventargs#changedrecipientfields)||
||[type](/javascript/api/outlook/outlook.recipientschangedeventargs#type)||
|[RecipientsChangedFields](/javascript/api/outlook/outlook.recipientschangedfields)|[bcc](/javascript/api/outlook/outlook.recipientschangedfields#bcc)|获取 " **密件抄送** " 字段中的收件人是否已更改。|
||[cc](/javascript/api/outlook/outlook.recipientschangedfields#cc)|获取是否更改了 **"抄送** " 字段中的收件人。|
||[optionalAttendees](/javascript/api/outlook/outlook.recipientschangedfields#optionalattendees)|获取可选的与会者是否已更改。|
||[requiredAttendees](/javascript/api/outlook/outlook.recipientschangedfields#requiredattendees)|获取所需与会者是否已更改。|
||[resources](/javascript/api/outlook/outlook.recipientschangedfields#resources)|获取是否更改了资源。|
||[to](/javascript/api/outlook/outlook.recipientschangedfields#to)|获取是否更改 **了 "收件人** " 字段中的收件人。|
|[定期](/javascript/api/outlook/outlook.recurrence)|[getAsync (选项？： AsyncContextOptions，回呼？： (asyncResult： <Recurrence>) => void) ](/javascript/api/outlook/outlook.recurrence#getasync-options--callback--asyncresult-)|返回约会系列的当前定期对象。|
||[recurrenceProperties](/javascript/api/outlook/outlook.recurrence#recurrenceproperties)|获取或设置定期约会系列的属性。|
||[recurrenceTimeZone](/javascript/api/outlook/outlook.recurrence#recurrencetimezone)|获取或设置定期约会系列的属性。|
||[recurrenceType](/javascript/api/outlook/outlook.recurrence#recurrencetype)|获取或设置定期约会系列的类型。|
||[seriesTime](/javascript/api/outlook/outlook.recurrence#seriestime)|{@Link SeriesTime | SeriesTime} 对象使您能够管理定期约会系列的开始和结束日期，以及|
||[setAsync (recurrencePattern：定期，选项？： AsyncContextOptions，回呼？： (asyncResult： <void>) => void) ](/javascript/api/outlook/outlook.recurrence#setasync-recurrencepattern--options--callback--asyncresult-)|设置约会系列的定期模式。|
|[RecurrenceChangedEventArgs](/javascript/api/outlook/outlook.recurrencechangedeventargs)|[循环](/javascript/api/outlook/outlook.recurrencechangedeventargs#recurrence)||
||[type](/javascript/api/outlook/outlook.recurrencechangedeventargs#type)||
|[RecurrenceProperties](/javascript/api/outlook/outlook.recurrenceproperties)|[dayOfMonth](/javascript/api/outlook/outlook.recurrenceproperties#dayofmonth)|表示月中的某一天。|
||[dayOfWeek](/javascript/api/outlook/outlook.recurrenceproperties#dayofweek)|表示一周中的某一天或一天的类型（例如，周末和工作日）。|
||[之前](/javascript/api/outlook/outlook.recurrenceproperties#days)|表示此定期事件的一组天数。|
||[firstDayOfWeek](/javascript/api/outlook/outlook.recurrenceproperties#firstdayofweek)|表示您选择的一周的第一天，否则默认值为当前用户设置中的值。|
||[interval](/javascript/api/outlook/outlook.recurrenceproperties#interval)|表示同一定期系列的两个实例之间的时间段。|
||[month](/javascript/api/outlook/outlook.recurrenceproperties#month)|表示月。|
||[weekNumber](/javascript/api/outlook/outlook.recurrenceproperties#weeknumber)|代表所选月份中的星期数，例如，月的第一周的 ' first '。|
|[RecurrenceTimeZone](/javascript/api/outlook/outlook.recurrencetimezone)|[name](/javascript/api/outlook/outlook.recurrencetimezone#name)|表示重复时区的名称。|
||[一定](/javascript/api/outlook/outlook.recurrencetimezone#offset)|整数值，表示在会议系列开始的日期之间的本地时区和 UTC 之间的差异（以分钟为单位）。|
|[SeriesTime](/javascript/api/outlook/outlook.seriestime)|[getDuration ( # B1 ](/javascript/api/outlook/outlook.seriestime#getduration--)|以分钟为单位获取定期约会系列中的常规实例的持续时间（以分钟为单位）。|
||[getEndDate ( # B1 ](/javascript/api/outlook/outlook.seriestime#getenddate--)|获取以下项中的定期模式的结束日期|
||[getEndTime ( # B1 ](/javascript/api/outlook/outlook.seriestime#getendtime--)|获取定期模式的常规约会或会议请求实例的结束时间（无论用户或在哪个时区中）|
||[getStartDate ( # B1 ](/javascript/api/outlook/outlook.seriestime#getstartdate--)|获取以下项中的定期模式的开始日期|
||[getStartTime ( # B1 ](/javascript/api/outlook/outlook.seriestime#getstarttime--)|获取定期模式的定期约会实例的开始时间（以用户/外接程序设置的任意时区）|
||[setDuration (分钟数： number) ](/javascript/api/outlook/outlook.seriestime#setduration-minutes-)|设置定期模式中所有约会的持续时间。|
||[setEndDate (日期： string) ](/javascript/api/outlook/outlook.seriestime#setenddate-date-)|设置定期约会系列的结束日期。|
||[setEndDate (year： number，month： number，day： number) ](/javascript/api/outlook/outlook.seriestime#setenddate-year--month--day-)|设置定期约会系列的结束日期。|
||[setStartDate (日期： string) ](/javascript/api/outlook/outlook.seriestime#setstartdate-date-)|设置定期约会系列的开始日期。|
||[setStartDate (year： number，month： number，day： number) ](/javascript/api/outlook/outlook.seriestime#setstartdate-year--month--day-)|设置定期约会系列的开始日期。|
||[setStartTime (小时数： number，分钟： number) ](/javascript/api/outlook/outlook.seriestime#setstarttime-hours--minutes-)|设置定期模式设置的任意时区的定期约会系列的所有实例的开始时间|
||[setStartTime (time： string) ](/javascript/api/outlook/outlook.seriestime#setstarttime-time-)|设置定期模式设置的任意时区的定期约会系列的所有实例的开始时间|
