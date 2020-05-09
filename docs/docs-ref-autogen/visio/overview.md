---
title: Office JavaScript API 参考
description: Office JavaScript Api 要求由主机进行设置。
ms.date: 05/05/2020
ms.openlocfilehash: 3a32c47b23fd6635c4c2b44b58ee9b351fffd8d5
ms.sourcegitcommit: 23d9a58660cb1dedf0bc414849a5aec519b419b3
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/07/2020
ms.locfileid: "44143928"
---
# <a name="office-javascript-api-reference"></a><span data-ttu-id="10b05-103">Office JavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="10b05-103">Office JavaScript API reference</span></span>

<span data-ttu-id="10b05-104">借助适用于 Office 的 JavaScript API，您可以创建可与 Office 主机应用程序中的对象模型进行交互的 Web 应用程序。</span><span class="sxs-lookup"><span data-stu-id="10b05-104">The JavaScript API for Office enables you to create web applications that interact with the object models in Office host applications.</span></span> <span data-ttu-id="10b05-105">使用此部分可详细了解可用于生成 Office 外接程序的类、方法和其他类型。</span><span class="sxs-lookup"><span data-stu-id="10b05-105">Use this section to learn more about the classes, methods, and other types available for building Office Add-ins.</span></span>

<span data-ttu-id="10b05-106">下面列出了主机特定的要求集（以及跨主机通用 Api）。</span><span class="sxs-lookup"><span data-stu-id="10b05-106">The following is a list of host-specific requirement sets (and the cross-host Common APIs).</span></span> <span data-ttu-id="10b05-107">每个项目都链接到该要求集支持的 API 参考文档版本（例如，ExcelApi 1.3 显示 ExcelApi 1.1、1.2、1.3 以及通用 API）的 api。</span><span class="sxs-lookup"><span data-stu-id="10b05-107">Each item links to a version of the API reference documentation that is supported by that requirement set (e.g. ExcelApi 1.3 shows APIs in ExcelApi 1.1, 1.2, 1.3 as well as the Common API).</span></span>

<span data-ttu-id="10b05-108">`ExcelApiOnline 1.1`是特殊要求集。</span><span class="sxs-lookup"><span data-stu-id="10b05-108">`ExcelApiOnline 1.1` is a special requirement set.</span></span> <span data-ttu-id="10b05-109">它包含适用于 web 上的 Excel 的最新 Api，但这些 Api 在所有平台中可能尚未完全受支持。</span><span class="sxs-lookup"><span data-stu-id="10b05-109">It contains the latest APIs for Excel on the web, but those APIs may not yet be fully supported across all platforms.</span></span> <span data-ttu-id="10b05-110">有关详细信息，请参阅[Excel JAVASCRIPT API online 要求集](/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set)。</span><span class="sxs-lookup"><span data-stu-id="10b05-110">See [Excel JavaScript API online-only requirement set](/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set) for more information.</span></span>

> [!TIP]
> <span data-ttu-id="10b05-111">选择此页面上的链接可查看指定要求集支持的 Api 的参考文档，或使用目录上方的筛选器选择下拉菜单更改要求集。</span><span class="sxs-lookup"><span data-stu-id="10b05-111">Choose a link on this page to view reference documentation for APIs supported by the specified requirement set, or use the filter selection drop-down menu above the table of contents to change the requirement set at any time.</span></span>

## <a name="excel"></a><span data-ttu-id="10b05-112">Excel</span><span class="sxs-lookup"><span data-stu-id="10b05-112">Excel</span></span>

- [<span data-ttu-id="10b05-113">ExcelApi 预览</span><span class="sxs-lookup"><span data-stu-id="10b05-113">ExcelApi Preview</span></span>](/javascript/api/excel?view=excel-js-preview)
- [<span data-ttu-id="10b05-114">ExcelApiOnline 1。1</span><span class="sxs-lookup"><span data-stu-id="10b05-114">ExcelApiOnline 1.1</span></span>](/javascript/api/excel?view=excel-js-online)
- [<span data-ttu-id="10b05-115">ExcelApi 1.11</span><span class="sxs-lookup"><span data-stu-id="10b05-115">ExcelApi 1.11</span></span>](/javascript/api/excel?view=excel-js-1.11)
- [<span data-ttu-id="10b05-116">ExcelApi 1.10</span><span class="sxs-lookup"><span data-stu-id="10b05-116">ExcelApi 1.10</span></span>](/javascript/api/excel?view=excel-js-1.10)
- [<span data-ttu-id="10b05-117">ExcelApi 1.9</span><span class="sxs-lookup"><span data-stu-id="10b05-117">ExcelApi 1.9</span></span>](/javascript/api/excel?view=excel-js-1.9)
- [<span data-ttu-id="10b05-118">ExcelApi 1.8</span><span class="sxs-lookup"><span data-stu-id="10b05-118">ExcelApi 1.8</span></span>](/javascript/api/excel?view=excel-js-1.8)
- [<span data-ttu-id="10b05-119">ExcelApi 1.7</span><span class="sxs-lookup"><span data-stu-id="10b05-119">ExcelApi 1.7</span></span>](/javascript/api/excel?view=excel-js-1.7)
- [<span data-ttu-id="10b05-120">ExcelApi 1.6</span><span class="sxs-lookup"><span data-stu-id="10b05-120">ExcelApi 1.6</span></span>](/javascript/api/excel?view=excel-js-1.6)
- [<span data-ttu-id="10b05-121">ExcelApi 1.5</span><span class="sxs-lookup"><span data-stu-id="10b05-121">ExcelApi 1.5</span></span>](/javascript/api/excel?view=excel-js-1.5)
- [<span data-ttu-id="10b05-122">ExcelApi 1.4</span><span class="sxs-lookup"><span data-stu-id="10b05-122">ExcelApi 1.4</span></span>](/javascript/api/excel?view=excel-js-1.4)
- [<span data-ttu-id="10b05-123">ExcelApi 1.3</span><span class="sxs-lookup"><span data-stu-id="10b05-123">ExcelApi 1.3</span></span>](/javascript/api/excel?view=excel-js-1.3)
- [<span data-ttu-id="10b05-124">ExcelApi 1.2</span><span class="sxs-lookup"><span data-stu-id="10b05-124">ExcelApi 1.2</span></span>](/javascript/api/excel?view=excel-js-1.2)
- [<span data-ttu-id="10b05-125">ExcelApi 1.1</span><span class="sxs-lookup"><span data-stu-id="10b05-125">ExcelApi 1.1</span></span>](/javascript/api/excel?view=excel-js-1.1)

## <a name="onenote"></a><span data-ttu-id="10b05-126">OneNote</span><span class="sxs-lookup"><span data-stu-id="10b05-126">OneNote</span></span>

- [<span data-ttu-id="10b05-127">OneNote 1。1</span><span class="sxs-lookup"><span data-stu-id="10b05-127">OneNote 1.1</span></span>](/javascript/api/onenote?view=onenote-js-1.1)

## <a name="outlook"></a><span data-ttu-id="10b05-128">Outlook</span><span class="sxs-lookup"><span data-stu-id="10b05-128">Outlook</span></span>

- [<span data-ttu-id="10b05-129">邮箱预览</span><span class="sxs-lookup"><span data-stu-id="10b05-129">Mailbox Preview</span></span>](/javascript/api/outlook?view=outlook-js-preview)
- [<span data-ttu-id="10b05-130">Mailbox 1.8</span><span class="sxs-lookup"><span data-stu-id="10b05-130">Mailbox 1.8</span></span>](/javascript/api/outlook?view=outlook-js-1.8)
- [<span data-ttu-id="10b05-131">Mailbox 1.7</span><span class="sxs-lookup"><span data-stu-id="10b05-131">Mailbox 1.7</span></span>](/javascript/api/outlook?view=outlook-js-1.7)
- [<span data-ttu-id="10b05-132">Mailbox 1.6</span><span class="sxs-lookup"><span data-stu-id="10b05-132">Mailbox 1.6</span></span>](/javascript/api/outlook?view=outlook-js-1.6)
- [<span data-ttu-id="10b05-133">Mailbox 1.5</span><span class="sxs-lookup"><span data-stu-id="10b05-133">Mailbox 1.5</span></span>](/javascript/api/outlook?view=outlook-js-1.5)
- [<span data-ttu-id="10b05-134">Mailbox 1.4</span><span class="sxs-lookup"><span data-stu-id="10b05-134">Mailbox 1.4</span></span>](/javascript/api/outlook?view=outlook-js-1.4)
- [<span data-ttu-id="10b05-135">Mailbox 1.3</span><span class="sxs-lookup"><span data-stu-id="10b05-135">Mailbox 1.3</span></span>](/javascript/api/outlook?view=outlook-js-1.3)
- [<span data-ttu-id="10b05-136">Mailbox 1.2</span><span class="sxs-lookup"><span data-stu-id="10b05-136">Mailbox 1.2</span></span>](/javascript/api/outlook?view=outlook-js-1.2)
- [<span data-ttu-id="10b05-137">Mailbox 1.1</span><span class="sxs-lookup"><span data-stu-id="10b05-137">Mailbox 1.1</span></span>](/javascript/api/outlook?view=outlook-js-1.1)

## <a name="powerpoint"></a><span data-ttu-id="10b05-138">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="10b05-138">PowerPoint</span></span>

- [<span data-ttu-id="10b05-139">PowerPointApi 1.1</span><span class="sxs-lookup"><span data-stu-id="10b05-139">PowerPointApi 1.1</span></span>](/javascript/api/powerpoint?view=powerpoint-js-1.1)

## <a name="visio"></a><span data-ttu-id="10b05-140">Visio</span><span class="sxs-lookup"><span data-stu-id="10b05-140">Visio</span></span>

- [<span data-ttu-id="10b05-141">VisioApi 1。1</span><span class="sxs-lookup"><span data-stu-id="10b05-141">VisioApi 1.1</span></span>](/javascript/api/visio?view=visio-js-1.1)

## <a name="word"></a><span data-ttu-id="10b05-142">Word</span><span class="sxs-lookup"><span data-stu-id="10b05-142">Word</span></span>

- [<span data-ttu-id="10b05-143">Word 预览</span><span class="sxs-lookup"><span data-stu-id="10b05-143">Word Preview</span></span>](/javascript/api/word?view=word-js-preview)
- [<span data-ttu-id="10b05-144">WordApi 1.3</span><span class="sxs-lookup"><span data-stu-id="10b05-144">WordApi 1.3</span></span>](/javascript/api/word?view=word-js-1.3)
- [<span data-ttu-id="10b05-145">WordApi 1.2</span><span class="sxs-lookup"><span data-stu-id="10b05-145">WordApi 1.2</span></span>](/javascript/api/word?view=word-js-1.2)
- [<span data-ttu-id="10b05-146">WordApi 1.1</span><span class="sxs-lookup"><span data-stu-id="10b05-146">WordApi 1.1</span></span>](/javascript/api/word?view=word-js-1.1)

## <a name="common-api"></a><span data-ttu-id="10b05-147">通用 API</span><span class="sxs-lookup"><span data-stu-id="10b05-147">Common API</span></span>

- [<span data-ttu-id="10b05-148">通用 API</span><span class="sxs-lookup"><span data-stu-id="10b05-148">Common API</span></span>](/javascript/api/office?view=common-js)

## <a name="see-also"></a><span data-ttu-id="10b05-149">另请参阅</span><span class="sxs-lookup"><span data-stu-id="10b05-149">See also</span></span>

- [<span data-ttu-id="10b05-150">关于 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="10b05-150">About Office Add-ins</span></span>](/office/dev/add-ins/overview)
- [<span data-ttu-id="10b05-151">Office 加载项主机和平台可用性</span><span class="sxs-lookup"><span data-stu-id="10b05-151">Office Add-in host and platform availability</span></span>](/office/dev/add-ins/overview/office-add-in-availability)
- [<span data-ttu-id="10b05-152">Office 版本和要求集</span><span class="sxs-lookup"><span data-stu-id="10b05-152">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="10b05-153">使用 Script Lab 探索 Office JavaScript API</span><span class="sxs-lookup"><span data-stu-id="10b05-153">Explore Office JavaScript API using Script Lab</span></span>](/office/dev/add-ins/overview/explore-with-script-lab)
