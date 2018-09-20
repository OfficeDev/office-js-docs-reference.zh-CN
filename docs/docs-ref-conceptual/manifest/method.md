# <a name="method-element"></a>Method 元素

指定来自适用于 Office 的 JavaScript API 的单个方法，Office 外接程序需要该方法才能激活。

**外接程序类型：** 内容、任务窗格

## <a name="syntax"></a>语法

```XML
<Method Name="string"/>
```

## <a name="contained-in"></a>包含在

[方法](methods.md)

## <a name="attributes"></a>Attributes

|**属性**|**类型**|**必需**|**说明**|
|:-----|:-----|:-----|:-----|
|名称|字符串|必需|指定由其父对象限定的所需方法的名称。例如，要指定 **getSelectedDataAsync** 方法，必须指定 `"Document.getSelectedDataAsync"`。|

## <a name="remarks"></a>说明

邮件外接程序不支持的**方法**和**方法**的元素。有关要求集的详细信息，请参阅[Office 版本和要求集](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)。

> [!IMPORTANT] 
> 由于无法指定单个方法的最低版本要求，以确保方法是在运行时，您应该还使用**if**语句的加载项在脚本中调用该方法时。 有关如何执行此操作的详细信息，请参阅[Understanding the JavaScript API for Office](https://docs.microsoft.com/office/dev/add-ins/develop/understanding-the-javascript-api-for-office)。

