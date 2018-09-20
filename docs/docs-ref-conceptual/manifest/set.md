# <a name="set-element"></a><span data-ttu-id="456be-101">Set 元素</span><span class="sxs-lookup"><span data-stu-id="456be-101">Set element</span></span>

<span data-ttu-id="456be-102">指定来自适用于 Office 的 JavaScript API 的要求集，Office 外接程序需要该集才能激活。</span><span class="sxs-lookup"><span data-stu-id="456be-102">Specifies a requirement set from the JavaScript API for Office that your Office Add-in requires to activate.</span></span>

<span data-ttu-id="456be-103">**外接程序类型：** 内容、任务窗格、邮件</span><span class="sxs-lookup"><span data-stu-id="456be-103">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="456be-104">语法</span><span class="sxs-lookup"><span data-stu-id="456be-104">Syntax</span></span>

```XML
<Set Name="string" MinVersion="n .n">
```

## <a name="contained-in"></a><span data-ttu-id="456be-105">包含在</span><span class="sxs-lookup"><span data-stu-id="456be-105">Contained in</span></span>

[<span data-ttu-id="456be-106">Sets</span><span class="sxs-lookup"><span data-stu-id="456be-106">Sets</span></span>](sets.md)

## <a name="attributes"></a><span data-ttu-id="456be-107">Attributes</span><span class="sxs-lookup"><span data-stu-id="456be-107">Attributes</span></span>

|<span data-ttu-id="456be-108">**属性**</span><span class="sxs-lookup"><span data-stu-id="456be-108">**Attribute**</span></span>|<span data-ttu-id="456be-109">**类型**</span><span class="sxs-lookup"><span data-stu-id="456be-109">**Type**</span></span>|<span data-ttu-id="456be-110">**必需**</span><span class="sxs-lookup"><span data-stu-id="456be-110">**Required**</span></span>|<span data-ttu-id="456be-111">**说明**</span><span class="sxs-lookup"><span data-stu-id="456be-111">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="456be-112">名称</span><span class="sxs-lookup"><span data-stu-id="456be-112">Name</span></span>|<span data-ttu-id="456be-113">字符串</span><span class="sxs-lookup"><span data-stu-id="456be-113">string</span></span>|<span data-ttu-id="456be-114">必需</span><span class="sxs-lookup"><span data-stu-id="456be-114">required</span></span>|<span data-ttu-id="456be-115">[要求集](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)名称。</span><span class="sxs-lookup"><span data-stu-id="456be-115">The name of a [requirement set](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>|
|<span data-ttu-id="456be-116">MinVersion</span><span class="sxs-lookup"><span data-stu-id="456be-116">MinVersion</span></span>|<span data-ttu-id="456be-117">字符串</span><span class="sxs-lookup"><span data-stu-id="456be-117">string</span></span>|<span data-ttu-id="456be-118">可选</span><span class="sxs-lookup"><span data-stu-id="456be-118">optional</span></span>|<span data-ttu-id="456be-p101">指定您的外接程序所需的 API 集的最低版本。如果 **DefaultMinVersion** 的值已在父 [Sets](sets.md) 元素中指定，则替代该值。</span><span class="sxs-lookup"><span data-stu-id="456be-p101">Specifies the minimum version of the API set required by your add-in. Overrides the value of  **DefaultMinVersion**, if it is specified in the parent [Sets](sets.md) element.</span></span>|

## <a name="remarks"></a><span data-ttu-id="456be-121">说明</span><span class="sxs-lookup"><span data-stu-id="456be-121">Remarks</span></span>

<span data-ttu-id="456be-122">有关要求集的详细信息，请参阅[Office 版本和要求集](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)。</span><span class="sxs-lookup"><span data-stu-id="456be-122">For more information about requirement sets, see [Office versions and requirement sets](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

<span data-ttu-id="456be-123">有关 **Set** 元素的 **MinVersion** 属性和 **Sets** 元素的 **DefaultMinVersion** 属性的详细信息，请参阅[在清单中设置 Requirements 元素](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest)。</span><span class="sxs-lookup"><span data-stu-id="456be-123">For more information about the  **MinVersion** attribute of the **Set** element and the **DefaultMinVersion** attribute of the **Sets** element, see [Set the Requirements element in the manifest](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="456be-124">对于邮件外接程序，没有只有一个`"Mailbox"`要求集可。</span><span class="sxs-lookup"><span data-stu-id="456be-124">For mail add-ins, there is only one  `"Mailbox"` requirement set available.</span></span> <span data-ttu-id="456be-125">此要求集包含整个 outlook 邮件加载项中支持的 API 子集，您必须指定`"Mailbox"`（不是可选的情况一样的内容和任务窗格的加载项） 设置邮件加载项的清单中的要求。</span><span class="sxs-lookup"><span data-stu-id="456be-125">This requirement set contains the entire subset of API supported in mail add-ins for Outlook, and you must specify the `"Mailbox"` requirement set in your mail add-in's manifest (it's not optional as is the case for content and task pane add-ins).</span></span> <span data-ttu-id="456be-126">此外，您无法声明邮件加载项中的特定方法的支持。</span><span class="sxs-lookup"><span data-stu-id="456be-126">Also, you can't declare support for specific methods in mail add-ins.</span></span>
