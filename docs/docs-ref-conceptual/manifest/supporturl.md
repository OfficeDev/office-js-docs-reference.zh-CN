# <a name="supporturl-element"></a><span data-ttu-id="89963-101">SupportUrl 元素</span><span class="sxs-lookup"><span data-stu-id="89963-101">SupportUrl element</span></span>

<span data-ttu-id="89963-102">指定提供外接程序支持信息的页面的 URL。</span><span class="sxs-lookup"><span data-stu-id="89963-102">Specifies the URL of a page that provides support information for your add-in.</span></span>

## <a name="syntax"></a><span data-ttu-id="89963-103">语法</span><span class="sxs-lookup"><span data-stu-id="89963-103">Syntax</span></span>

```XML
<OfficeApp>
...
  <IconUrl DefaultValue="https://contoso.com/assets/icon-32.png" />
  <HighResolutionIconUrl DefaultValue="https://contoso.com/assets/hi-res-icon.png"/>
  
  
  <SupportUrl DefaultValue="https://contoso.com/support " />
  
  
  <AppDomains>
  ...
  </AppDomains>
...
</OfficeApp>
```

## <a name="contained-in"></a><span data-ttu-id="89963-104">包含在</span><span class="sxs-lookup"><span data-stu-id="89963-104">Contained in</span></span>

[<span data-ttu-id="89963-105">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="89963-105">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="89963-106">可以包含</span><span class="sxs-lookup"><span data-stu-id="89963-106">Can contain</span></span>

|  <span data-ttu-id="89963-107">元素</span><span class="sxs-lookup"><span data-stu-id="89963-107">Element</span></span> | <span data-ttu-id="89963-108">必需</span><span class="sxs-lookup"><span data-stu-id="89963-108">Required</span></span> | <span data-ttu-id="89963-109">说明</span><span class="sxs-lookup"><span data-stu-id="89963-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="89963-110">Override</span><span class="sxs-lookup"><span data-stu-id="89963-110">Override</span></span>](override.md)   | <span data-ttu-id="89963-111">否</span><span class="sxs-lookup"><span data-stu-id="89963-111">No</span></span> | <span data-ttu-id="89963-112">指定其他区域设置 URL 的设置</span><span class="sxs-lookup"><span data-stu-id="89963-112">Specifies the setting for additional locale urls</span></span> |

## <a name="attributes"></a><span data-ttu-id="89963-113">Attributes</span><span class="sxs-lookup"><span data-stu-id="89963-113">Attributes</span></span>

|<span data-ttu-id="89963-114">**属性**</span><span class="sxs-lookup"><span data-stu-id="89963-114">**Attribute**</span></span>|<span data-ttu-id="89963-115">**类型**</span><span class="sxs-lookup"><span data-stu-id="89963-115">**Type**</span></span>|<span data-ttu-id="89963-116">**必需**</span><span class="sxs-lookup"><span data-stu-id="89963-116">**Required**</span></span>|<span data-ttu-id="89963-117">**说明**</span><span class="sxs-lookup"><span data-stu-id="89963-117">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="89963-118">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="89963-118">DefaultValue</span></span>|<span data-ttu-id="89963-119">URL</span><span class="sxs-lookup"><span data-stu-id="89963-119">URL</span></span>|<span data-ttu-id="89963-120">必需</span><span class="sxs-lookup"><span data-stu-id="89963-120">required</span></span>|<span data-ttu-id="89963-121">指定该设置的默认值，表示为 [DefaultLocale](defaultlocale.md) 元素中指定的区域设置。</span><span class="sxs-lookup"><span data-stu-id="89963-121">Specifies the default value for this setting, expressed for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|
