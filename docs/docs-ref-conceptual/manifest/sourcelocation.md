# <a name="sourcelocation-element"></a>SourceLocation 元素

指定 Office 外接程序的源文件位置为介于 1 和 2018 个字符之间的 URL。源位置必须是 HTTPS 地址，而非文件路径。

**外接程序类型：** 内容、任务窗格、邮件

## <a name="syntax"></a>语法

```XML
<SourceLocation DefaultValue="string" />
```

## <a name="contained-in"></a>包含在

- [DefaultSettings](defaultsettings.md)（内容和任务窗格外接程序）
- [FormSettings](formsettings.md)（邮件外接程序）
- [ExtensionPoint](extensionpoint.md)（上下文邮件外接程序）

## <a name="can-contain"></a>可以包含

[Override](override.md)

## <a name="attributes"></a>属性

|**属性**|**类型**|**必需**|**说明**|
|:-----|:-----|:-----|:-----|
|DefaultValue|URL|必需|指定该设置的默认值，表示为 [DefaultLocale](defaultlocale.md) 元素中指定的区域设置。|
