# <a name="officeapp-element"></a>OfficeApp 元素

Office 外接程序清单中的根元素。

**外接程序类型：** 内容、任务窗格、邮件

## <a name="syntax"></a>语法

```XML
<OfficeApp 
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
  xsi:type= ["ContentApp" |"MailApp"| "TaskPaneApp"]>
  ...
</OfficeApp>
```

## <a name="contained-in"></a>包含在

 _none_

## <a name="must-contain"></a>必须包含

|**元素**|**内容**|**邮件**|**TaskPane**|
|:-----|:-----|:-----|:-----|
|
  [Id](id.md)|x|x|x|
|[Version](version.md)|x|x|x|
|[ProviderName](providername.md)|x|x|x|
|[DefaultLocale](defaultlocale.md)|x|x|x|
|[DefaultSettings](defaultsettings.md)|x||x|
|[DisplayName](displayname.md)|x|x|x|
|[说明](description.md)|x|x|x|
|[FormSettings](formsettings.md)||x||
|[权限](permissions.md)|x||x|
|[Rule](rule.md)||x||

## <a name="can-contain"></a>可以包含

|**元素**|**内容**|**邮件**|**TaskPane**|
|:-----|:-----|:-----|:-----|
|[AlternateId](alternateid.md)|x|x|x|
|[IconUrl](iconurl.md)|x|x|x|
|[HighResolutionIconUrl](highresolutioniconurl.md)|x|x|x|
|[SupportUrl](supporturl.md)|x|x|x|
|[AppDomains](appdomains.md)|x|x|x|
|[主机](hosts.md)|x|x|x|
|[要求](requirements.md)|x|x|x|
|[AllowSnapshot](allowsnapshot.md)|x|||
|[权限](permissions.md)||x||
|[DisableEntityHighlighting](disableentityhighlighting.md)||x||
|"词典"|||x|
|[VersionOverrides](versionoverrides.md)|X|X|X|

## <a name="attributes"></a>属性

|||
|:-----|:-----|
|xmlns|定义的 Office 外接程序清单命名空间和架构版本。应始终将此属性设置为 `"http://schemas.microsoft.com/office/appforoffice/1.1"`|
|xmlns:xsi|定义 XMLSchema 实例。应始终将此属性设置为 `"http://www.w3.org/2001/XMLSchema-instance"`|
|xsi:type|定义的 Office 外接程序的类型。应始终将此属性设置为下列值之一：`"ContentApp"`、`"MailApp"` 或 `"TaskPaneApp"`|
