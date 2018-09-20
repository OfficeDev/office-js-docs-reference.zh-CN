# <a name="override-element"></a><span data-ttu-id="bbbd2-101">Override 元素</span><span class="sxs-lookup"><span data-stu-id="bbbd2-101">Override element</span></span>

<span data-ttu-id="bbbd2-102">提供一种为其他区域设置指定某设置的值的方法。</span><span class="sxs-lookup"><span data-stu-id="bbbd2-102">Provides a way to specify the value of a setting for an additional locale.</span></span>

<span data-ttu-id="bbbd2-103">**外接程序类型：** 内容、任务窗格、邮件</span><span class="sxs-lookup"><span data-stu-id="bbbd2-103">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="bbbd2-104">语法</span><span class="sxs-lookup"><span data-stu-id="bbbd2-104">Syntax</span></span>

```XML
<Override Locale="string" Value="string" />
```

## <a name="contained-in"></a><span data-ttu-id="bbbd2-105">包含在</span><span class="sxs-lookup"><span data-stu-id="bbbd2-105">Contained in</span></span>

|<span data-ttu-id="bbbd2-106">**元素**</span><span class="sxs-lookup"><span data-stu-id="bbbd2-106">**Element**</span></span>|
|:-----|
|[<span data-ttu-id="bbbd2-107">CitationText</span><span class="sxs-lookup"><span data-stu-id="bbbd2-107">CitationText</span></span>](citationtext.md)|
|[<span data-ttu-id="bbbd2-108">说明</span><span class="sxs-lookup"><span data-stu-id="bbbd2-108">Description</span></span>](description.md)|
|[<span data-ttu-id="bbbd2-109">DictionaryName</span><span class="sxs-lookup"><span data-stu-id="bbbd2-109">DictionaryName</span></span>](dictionaryname.md)|
|[<span data-ttu-id="bbbd2-110">DictionaryHomePage</span><span class="sxs-lookup"><span data-stu-id="bbbd2-110">DictionaryHomePage</span></span>](dictionaryhomepage.md)|
|[<span data-ttu-id="bbbd2-111">DisplayName</span><span class="sxs-lookup"><span data-stu-id="bbbd2-111">DisplayName</span></span>](displayname.md)|
|[<span data-ttu-id="bbbd2-112">HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="bbbd2-112">HighResolutionIconUrl</span></span>](highresolutioniconurl.md)|
|[<span data-ttu-id="bbbd2-113">IconUrl</span><span class="sxs-lookup"><span data-stu-id="bbbd2-113">IconUrl</span></span>](iconurl.md)|
|[<span data-ttu-id="bbbd2-114">QueryUri</span><span class="sxs-lookup"><span data-stu-id="bbbd2-114">QueryUri</span></span>](queryuri.md)|
|[<span data-ttu-id="bbbd2-115">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="bbbd2-115">SourceLocation</span></span>](sourcelocation.md)|
|[<span data-ttu-id="bbbd2-116">SupportUrl</span><span class="sxs-lookup"><span data-stu-id="bbbd2-116">SupportUrl</span></span>](supporturl.md)|

## <a name="attributes"></a><span data-ttu-id="bbbd2-117">属性</span><span class="sxs-lookup"><span data-stu-id="bbbd2-117">Attributes</span></span>

|<span data-ttu-id="bbbd2-118">**属性**</span><span class="sxs-lookup"><span data-stu-id="bbbd2-118">**Attribute**</span></span>|<span data-ttu-id="bbbd2-119">**类型**</span><span class="sxs-lookup"><span data-stu-id="bbbd2-119">**Type**</span></span>|<span data-ttu-id="bbbd2-120">**必需**</span><span class="sxs-lookup"><span data-stu-id="bbbd2-120">**Required**</span></span>|<span data-ttu-id="bbbd2-121">**说明**</span><span class="sxs-lookup"><span data-stu-id="bbbd2-121">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="bbbd2-122">Locale</span><span class="sxs-lookup"><span data-stu-id="bbbd2-122">Locale</span></span>|<span data-ttu-id="bbbd2-123">字符串</span><span class="sxs-lookup"><span data-stu-id="bbbd2-123">string</span></span>|<span data-ttu-id="bbbd2-124">必需</span><span class="sxs-lookup"><span data-stu-id="bbbd2-124">required</span></span>|<span data-ttu-id="bbbd2-125">为此替代项指定区域设置的区域性名称，采用 BCP 47 语言标记格式，例如 `"en-US"`。</span><span class="sxs-lookup"><span data-stu-id="bbbd2-125">Specifies the culture name of the locale for this override in the BCP 47 language tag format, such as  `"en-US"`.</span></span>|
|<span data-ttu-id="bbbd2-126">值</span><span class="sxs-lookup"><span data-stu-id="bbbd2-126">Value</span></span>|<span data-ttu-id="bbbd2-127">字符串</span><span class="sxs-lookup"><span data-stu-id="bbbd2-127">string</span></span>|<span data-ttu-id="bbbd2-128">必需</span><span class="sxs-lookup"><span data-stu-id="bbbd2-128">required</span></span>|<span data-ttu-id="bbbd2-129">指定表示为指定区域设置的设置的值。</span><span class="sxs-lookup"><span data-stu-id="bbbd2-129">Specifies value of the setting expressed for the specified locale.</span></span>|

## <a name="see-also"></a><span data-ttu-id="bbbd2-130">另请参阅</span><span class="sxs-lookup"><span data-stu-id="bbbd2-130">See also</span></span>

- [<span data-ttu-id="bbbd2-131">Office 外接程序的本地化</span><span class="sxs-lookup"><span data-stu-id="bbbd2-131">Localization for Office Add-ins</span></span>](https://docs.microsoft.com/office/dev/add-ins/develop/localization)
    
