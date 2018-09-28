# <a name="override-element"></a><span data-ttu-id="57244-101">Override 元素</span><span class="sxs-lookup"><span data-stu-id="57244-101">Override element</span></span>

<span data-ttu-id="57244-102">提供一种为其他区域设置指定某设置的值的方法。</span><span class="sxs-lookup"><span data-stu-id="57244-102">Provides a way to specify the value of a setting for an additional locale.</span></span>

<span data-ttu-id="57244-103">**外接程序类型：** 内容、任务窗格、邮件</span><span class="sxs-lookup"><span data-stu-id="57244-103">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="57244-104">语法</span><span class="sxs-lookup"><span data-stu-id="57244-104">Syntax</span></span>

```XML
<Override Locale="string" Value="string" />
```

## <a name="contained-in"></a><span data-ttu-id="57244-105">包含在</span><span class="sxs-lookup"><span data-stu-id="57244-105">Contained in</span></span>

|<span data-ttu-id="57244-106">**元素**</span><span class="sxs-lookup"><span data-stu-id="57244-106">**Element**</span></span>|
|:-----|
|[<span data-ttu-id="57244-107">CitationText</span><span class="sxs-lookup"><span data-stu-id="57244-107">CitationText</span></span>](citationtext.md)|
|[<span data-ttu-id="57244-108">说明</span><span class="sxs-lookup"><span data-stu-id="57244-108">Description</span></span>](description.md)|
|[<span data-ttu-id="57244-109">DictionaryName</span><span class="sxs-lookup"><span data-stu-id="57244-109">DictionaryName</span></span>](dictionaryname.md)|
|[<span data-ttu-id="57244-110">DictionaryHomePage</span><span class="sxs-lookup"><span data-stu-id="57244-110">DictionaryHomePage</span></span>](dictionaryhomepage.md)|
|[<span data-ttu-id="57244-111">DisplayName</span><span class="sxs-lookup"><span data-stu-id="57244-111">DisplayName</span></span>](displayname.md)|
|[<span data-ttu-id="57244-112">HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="57244-112">HighResolutionIconUrl</span></span>](highresolutioniconurl.md)|
|[<span data-ttu-id="57244-113">IconUrl</span><span class="sxs-lookup"><span data-stu-id="57244-113">IconUrl</span></span>](iconurl.md)|
|[<span data-ttu-id="57244-114">QueryUri</span><span class="sxs-lookup"><span data-stu-id="57244-114">QueryUri</span></span>](queryuri.md)|
|[<span data-ttu-id="57244-115">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="57244-115">SourceLocation</span></span>](sourcelocation.md)|
|[<span data-ttu-id="57244-116">SupportUrl</span><span class="sxs-lookup"><span data-stu-id="57244-116">SupportUrl</span></span>](supporturl.md)|

## <a name="attributes"></a><span data-ttu-id="57244-117">属性</span><span class="sxs-lookup"><span data-stu-id="57244-117">Attributes</span></span>

|<span data-ttu-id="57244-118">**属性**</span><span class="sxs-lookup"><span data-stu-id="57244-118">**Attribute**</span></span>|<span data-ttu-id="57244-119">**类型**</span><span class="sxs-lookup"><span data-stu-id="57244-119">**Type**</span></span>|<span data-ttu-id="57244-120">**必需**</span><span class="sxs-lookup"><span data-stu-id="57244-120">**Required**</span></span>|<span data-ttu-id="57244-121">**说明**</span><span class="sxs-lookup"><span data-stu-id="57244-121">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="57244-122">Locale</span><span class="sxs-lookup"><span data-stu-id="57244-122">Locale</span></span>|<span data-ttu-id="57244-123">字符串</span><span class="sxs-lookup"><span data-stu-id="57244-123">string</span></span>|<span data-ttu-id="57244-124">必需</span><span class="sxs-lookup"><span data-stu-id="57244-124">required</span></span>|<span data-ttu-id="57244-125">为此替代项指定区域设置的区域性名称，采用 BCP 47 语言标记格式，例如 `"en-US"`。</span><span class="sxs-lookup"><span data-stu-id="57244-125">Specifies the culture name of the locale for this override in the BCP 47 language tag format, such as  `"en-US"`.</span></span>|
|<span data-ttu-id="57244-126">值</span><span class="sxs-lookup"><span data-stu-id="57244-126">Value</span></span>|<span data-ttu-id="57244-127">字符串</span><span class="sxs-lookup"><span data-stu-id="57244-127">string</span></span>|<span data-ttu-id="57244-128">必需</span><span class="sxs-lookup"><span data-stu-id="57244-128">required</span></span>|<span data-ttu-id="57244-129">指定表示为指定区域设置的设置的值。</span><span class="sxs-lookup"><span data-stu-id="57244-129">Specifies value of the setting expressed for the specified locale.</span></span>|

## <a name="see-also"></a><span data-ttu-id="57244-130">另请参阅</span><span class="sxs-lookup"><span data-stu-id="57244-130">See also</span></span>

- [<span data-ttu-id="57244-131">Office 外接程序的本地化</span><span class="sxs-lookup"><span data-stu-id="57244-131">Localization for Office Add-ins</span></span>](https://docs.microsoft.com/office/dev/add-ins/develop/localization)
    
