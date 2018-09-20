# <a name="override-element"></a>Override 元素

提供一种为其他区域设置指定某设置的值的方法。

**外接程序类型：** 内容、任务窗格、邮件

## <a name="syntax"></a>语法

```XML
<Override Locale="string" Value="string" />
```

## <a name="contained-in"></a>包含在

|**元素**|
|:-----|
|[CitationText](citationtext.md)|
|[说明](description.md)|
|[DictionaryName](dictionaryname.md)|
|[DictionaryHomePage](dictionaryhomepage.md)|
|[DisplayName](displayname.md)|
|[HighResolutionIconUrl](highresolutioniconurl.md)|
|[IconUrl](iconurl.md)|
|[QueryUri](queryuri.md)|
|[SourceLocation](sourcelocation.md)|
|[SupportUrl](supporturl.md)|

## <a name="attributes"></a>属性

|**属性**|**类型**|**必需**|**说明**|
|:-----|:-----|:-----|:-----|
|Locale|字符串|必需|为此替代项指定区域设置的区域性名称，采用 BCP 47 语言标记格式，例如 `"en-US"`。|
|值|字符串|必需|指定表示为指定区域设置的设置的值。|

## <a name="see-also"></a>另请参阅

- [Office 外接程序的本地化](https://docs.microsoft.com/office/dev/add-ins/develop/localization)
    
