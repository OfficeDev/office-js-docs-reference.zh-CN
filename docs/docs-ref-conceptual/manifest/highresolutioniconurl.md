# <a name="highresolutioniconurl-element"></a><span data-ttu-id="d78d1-101">HighResolutionIconUrl 元素</span><span class="sxs-lookup"><span data-stu-id="d78d1-101">HighResolutionIconUrl element</span></span>

<span data-ttu-id="d78d1-102">指定用于表示插入 UX 中的 Office 外接程序和高 DPI 屏幕上的 Office 应用商店的图像的 URL。</span><span class="sxs-lookup"><span data-stu-id="d78d1-102">Specifies the URL of the image that is used to represent your Office Add-in in the insertion UX and Office Store on high DPI screens.</span></span>

<span data-ttu-id="d78d1-103">**外接程序类型：** 内容、任务窗格、邮件</span><span class="sxs-lookup"><span data-stu-id="d78d1-103">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="d78d1-104">语法</span><span class="sxs-lookup"><span data-stu-id="d78d1-104">Syntax</span></span>

```XML
<HighResolutionIconUrl DefaultValue="string" />
```

## <a name="can-contain"></a><span data-ttu-id="d78d1-105">可以包含</span><span class="sxs-lookup"><span data-stu-id="d78d1-105">Can contain</span></span>

[<span data-ttu-id="d78d1-106">Override</span><span class="sxs-lookup"><span data-stu-id="d78d1-106">Override</span></span>](override.md)

## <a name="attributes"></a><span data-ttu-id="d78d1-107">属性</span><span class="sxs-lookup"><span data-stu-id="d78d1-107">Attributes</span></span>

|<span data-ttu-id="d78d1-108">**属性**</span><span class="sxs-lookup"><span data-stu-id="d78d1-108">**Attribute**</span></span>|<span data-ttu-id="d78d1-109">**类型**</span><span class="sxs-lookup"><span data-stu-id="d78d1-109">**Type**</span></span>|<span data-ttu-id="d78d1-110">**必需**</span><span class="sxs-lookup"><span data-stu-id="d78d1-110">**Required**</span></span>|<span data-ttu-id="d78d1-111">**说明**</span><span class="sxs-lookup"><span data-stu-id="d78d1-111">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="d78d1-112">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="d78d1-112">DefaultValue</span></span>|<span data-ttu-id="d78d1-113">字符串 (URL)</span><span class="sxs-lookup"><span data-stu-id="d78d1-113">string (URL)</span></span>|<span data-ttu-id="d78d1-114">必需</span><span class="sxs-lookup"><span data-stu-id="d78d1-114">required</span></span>|<span data-ttu-id="d78d1-115">指定该设置的默认值，表示为 [DefaultLocale](defaultlocale.md) 元素中指定的区域设置。</span><span class="sxs-lookup"><span data-stu-id="d78d1-115">Specifies the default value for this setting, expressed for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|

## <a name="remarks"></a><span data-ttu-id="d78d1-116">说明</span><span class="sxs-lookup"><span data-stu-id="d78d1-116">Remarks</span></span>

<span data-ttu-id="d78d1-p101">对于邮件外接程序，图标显示在“**文件**” > “**管理外接程序**”UI 中。对于内容或任务窗格外接程序，图标显示在“**插入**” > “**外接程序**”UI 中。</span><span class="sxs-lookup"><span data-stu-id="d78d1-p101">For a mail add-in, the icon is displayed in the  **File** > **Manage add-ins** UI . For a content or task pane add-in, the icon is displayed in the **Insert** > **Add-ins** UI.</span></span>

<span data-ttu-id="d78d1-119">图像必须以 64 x 64 像素的分辨率采用下列任一文件格式进行保存：GIF、JPG、PNG、EXIF、BMP 或 TIFF。</span><span class="sxs-lookup"><span data-stu-id="d78d1-119">The image must be in one of the following file formats at a recommended resolution of 64 x 64 pixels: GIF, JPG, PNG, EXIF, BMP or TIFF.</span></span> <span data-ttu-id="d78d1-120">有关详细信息，请参阅[创建有效列表在 AppSource 中，在 Office](https://docs.microsoft.com/office/dev/store/create-effective-office-store-listings)中_创建一致的视觉标识您的应用程序_一节。</span><span class="sxs-lookup"><span data-stu-id="d78d1-120">For more information, see the section  _Create a consistent visual identity for your app_ in [Create effective listings in AppSource and within Office](https://docs.microsoft.com/office/dev/store/create-effective-office-store-listings).</span></span>
