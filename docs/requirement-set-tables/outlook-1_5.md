| Class | 域 | 说明 |
|:---|:---|:---|
|[邮箱](/javascript/api/outlook/outlook.mailbox)|[addHandlerAsync (事件类型： AsyncContextOptions \| 的字符串、处理程序： (类型：) => void、选项？：、callback？： (asyncResult： <void>) => void) ](/javascript/api/outlook/outlook.mailbox#addhandlerasync-eventtype--handler--type-)|添加支持事件的事件处理程序。|
||[Mailbox.getcallbacktokenasync (选项： AsyncContextOptions & {isRest？： boolean}、callback： (asyncResult： <string>) => void) ](/javascript/api/outlook/outlook.mailbox#getcallbacktokenasync-options--isrest--callback--asyncresult-)|获取一个字符串，其中包含用于调用 REST Api 或 Exchange Web Services (EWS) 的标记。|
||[isRest](/javascript/api/outlook/outlook.mailbox#isrest)||
||[removeHandlerAsync (事件类型： Office. 命令 \| 行字符串、选项？： AsyncContextOptions、callback？： (asyncResult： <void>) => void) ](/javascript/api/outlook/outlook.mailbox#removehandlerasync-eventtype--options--callback--asyncresult-)|删除受支持事件类型的事件处理程序。|
||[restUrl](/javascript/api/outlook/outlook.mailbox#resturl)|获取此电子邮件帐户的 REST 终结点的 URL。|
