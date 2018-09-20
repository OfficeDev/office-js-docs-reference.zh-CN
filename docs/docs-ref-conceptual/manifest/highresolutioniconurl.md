# <a name="highresolutioniconurl-element"></a>HighResolutionIconUrl 元素

指定用于表示插入 UX 中的 Office 外接程序和高 DPI 屏幕上的 Office 应用商店的图像的 URL。

**外接程序类型：** 内容、任务窗格、邮件

## <a name="syntax"></a>语法

```XML
<HighResolutionIconUrl DefaultValue="string" />
```

## <a name="can-contain"></a>可以包含

[Override](override.md)

## <a name="attributes"></a>属性

|**属性**|**类型**|**必需**|**说明**|
|:-----|:-----|:-----|:-----|
|DefaultValue|字符串 (URL)|必需|指定该设置的默认值，表示为 [DefaultLocale](defaultlocale.md) 元素中指定的区域设置。|

## <a name="remarks"></a>说明

对于邮件外接程序，图标显示在“**文件**” > “**管理外接程序**”UI 中。对于内容或任务窗格外接程序，图标显示在“**插入**” > “**外接程序**”UI 中。

图像必须以 64 x 64 像素的分辨率采用下列任一文件格式进行保存：GIF、JPG、PNG、EXIF、BMP 或 TIFF。 有关详细信息，请参阅[创建有效列表在 AppSource 中，在 Office](https://docs.microsoft.com/office/dev/store/create-effective-office-store-listings)中_创建一致的视觉标识您的应用程序_一节。
