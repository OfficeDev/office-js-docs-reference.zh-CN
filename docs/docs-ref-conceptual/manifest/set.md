# <a name="set-element"></a>Set 元素

指定来自适用于 Office 的 JavaScript API 的要求集，Office 外接程序需要该集才能激活。

**外接程序类型：** 内容、任务窗格、邮件

## <a name="syntax"></a>语法

```XML
<Set Name="string" MinVersion="n .n">
```

## <a name="contained-in"></a>包含在

[Sets](sets.md)

## <a name="attributes"></a>Attributes

|**属性**|**类型**|**必需**|**说明**|
|:-----|:-----|:-----|:-----|
|名称|字符串|必需|[要求集](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)名称。|
|MinVersion|字符串|可选|指定您的外接程序所需的 API 集的最低版本。如果 **DefaultMinVersion** 的值已在父 [Sets](sets.md) 元素中指定，则替代该值。|

## <a name="remarks"></a>说明

有关要求集的详细信息，请参阅[Office 版本和要求集](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)。

有关 **Set** 元素的 **MinVersion** 属性和 **Sets** 元素的 **DefaultMinVersion** 属性的详细信息，请参阅[在清单中设置 Requirements 元素](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest)。

> [!IMPORTANT] 
> 对于邮件外接程序，没有只有一个`"Mailbox"`要求集可。 此要求集包含整个 outlook 邮件加载项中支持的 API 子集，您必须指定`"Mailbox"`（不是可选的情况一样的内容和任务窗格的加载项） 设置邮件加载项的清单中的要求。 此外，您无法声明邮件加载项中的特定方法的支持。
