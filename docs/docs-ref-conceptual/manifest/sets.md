# <a name="sets-element"></a><span data-ttu-id="cd31f-101">Sets 元素</span><span class="sxs-lookup"><span data-stu-id="cd31f-101">Sets element</span></span>

<span data-ttu-id="cd31f-102">指定适用于 Office 的 JavaScript API 的最小子集，Office 外接程序需要该子集才能激活。</span><span class="sxs-lookup"><span data-stu-id="cd31f-102">Specifies the minimum subset of the JavaScript API for Office that your Office Add-in requires in order to activate.</span></span>

<span data-ttu-id="cd31f-103">**外接程序类型：** 内容、任务窗格、邮件</span><span class="sxs-lookup"><span data-stu-id="cd31f-103">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="cd31f-104">语法</span><span class="sxs-lookup"><span data-stu-id="cd31f-104">Syntax</span></span>

```XML
<Sets DefaultMinVersion="n .n ">
   ...
</Sets>
```

## <a name="contained-in"></a><span data-ttu-id="cd31f-105">包含在</span><span class="sxs-lookup"><span data-stu-id="cd31f-105">Contained in</span></span>

[<span data-ttu-id="cd31f-106">要求</span><span class="sxs-lookup"><span data-stu-id="cd31f-106">Requirements</span></span>](requirements.md)

## <a name="can-contain"></a><span data-ttu-id="cd31f-107">可以包含</span><span class="sxs-lookup"><span data-stu-id="cd31f-107">Can contain</span></span>

[<span data-ttu-id="cd31f-108">Set</span><span class="sxs-lookup"><span data-stu-id="cd31f-108">Set</span></span>](set.md)

## <a name="attributes"></a><span data-ttu-id="cd31f-109">Attributes</span><span class="sxs-lookup"><span data-stu-id="cd31f-109">Attributes</span></span>

|<span data-ttu-id="cd31f-110">**属性**</span><span class="sxs-lookup"><span data-stu-id="cd31f-110">**Attribute**</span></span>|<span data-ttu-id="cd31f-111">**类型**</span><span class="sxs-lookup"><span data-stu-id="cd31f-111">**Type**</span></span>|<span data-ttu-id="cd31f-112">**必需**</span><span class="sxs-lookup"><span data-stu-id="cd31f-112">**Required**</span></span>|<span data-ttu-id="cd31f-113">**说明**</span><span class="sxs-lookup"><span data-stu-id="cd31f-113">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="cd31f-114">DefaultMinVersion</span><span class="sxs-lookup"><span data-stu-id="cd31f-114">DefaultMinVersion</span></span>|<span data-ttu-id="cd31f-115">字符串</span><span class="sxs-lookup"><span data-stu-id="cd31f-115">string</span></span>|<span data-ttu-id="cd31f-116">可选</span><span class="sxs-lookup"><span data-stu-id="cd31f-116">optional</span></span>|<span data-ttu-id="cd31f-p101">为所有子 **Set** 元素指定默认的 [MinVersion](set.md) 属性值。默认值为“1.1”。</span><span class="sxs-lookup"><span data-stu-id="cd31f-p101">Specifies the default  **MinVersion** attribute value for all child [Set](set.md) elements. The default value is "1.1".</span></span>|

## <a name="remarks"></a><span data-ttu-id="cd31f-119">说明</span><span class="sxs-lookup"><span data-stu-id="cd31f-119">Remarks</span></span>

<span data-ttu-id="cd31f-120">有关要求集的详细信息，请参阅[Office 版本和要求集](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)。</span><span class="sxs-lookup"><span data-stu-id="cd31f-120">For more information about requirement sets, see [Office versions and requirement sets](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

<span data-ttu-id="cd31f-121">有关 **Set** 元素的 **MinVersion** 属性和 **Sets** 元素的 **DefaultMinVersion** 属性的详细信息，请参阅[在清单中设置 Requirements 元素](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest)。</span><span class="sxs-lookup"><span data-stu-id="cd31f-121">For more information about the  **MinVersion** attribute of the **Set** element and the **DefaultMinVersion** attribute of the **Sets** element, see [Set the Requirements element in the manifest](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).</span></span>

