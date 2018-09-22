# <a name="outlook-add-in-api-requirement-set-16"></a>Outlook 加载项 API 要求集 1.6

Outlook 加载项 API 子集的 JavaScript API for Office 包括对象、 方法、 属性和事件，您可以在 Outlook 中使用外接程序。

## <a name="whats-new-in-16"></a>What's new in 1.6？

要求集 1.6 包括所有[要求设置 1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md)的功能。 它添加以下功能。

- 添加了新的 Api 的上下文的加载项获取实体或正则表达式匹配用户选择要激活的加载项。
- 添加一个新的 API，以打开新邮件窗体。
- 添加的外接程序来确定用户的邮箱的帐户类型的功能。

### <a name="change-log"></a>更改日志

- 添加[Office.context.mailbox.item.getSelectedEntities](office.context.mailbox.item.md#getselectedentities--entitiesjavascriptapioutlook16officeentities)： 添加新的函数获取用户选择突出显示匹配项中找到的实体。 突出显示的匹配项适用于上下文加载项。
- 添加[Office.context.mailbox.item.getSelectedRegExMatches](office.context.mailbox.item.md#getselectedregexmatches--object)： 添加新函数的返回中突出显示匹配项匹配在清单 XML 文件中定义的正则表达式的字符串值。 突出显示的匹配项适用于上下文加载项。
- 添加[Office.context.mailbox.displayNewMessageForm](office.context.mailbox.md#displaynewmessageformparameters)： 添加新函数的打开一个新的邮件窗体。
- 添加[Office.context.mailbox.userProfile.accountType](office.context.mailbox.userprofile.md#accounttype-string)： 将新的成员添加到用户配置文件，指示用户的帐户的类型。

## <a name="see-also"></a>另请参阅

- [Outlook 外接程序](https://docs.microsoft.com/outlook/add-ins/)
- [Outlook 外接程序代码示例](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [入门](https://docs.microsoft.com/outlook/add-ins/quick-start)