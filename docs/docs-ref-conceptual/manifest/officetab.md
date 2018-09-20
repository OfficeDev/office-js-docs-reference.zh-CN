# <a name="officetab-element"></a><span data-ttu-id="13bc9-101">OfficeTab 元素</span><span class="sxs-lookup"><span data-stu-id="13bc9-101">OfficeTab element</span></span>

<span data-ttu-id="13bc9-p101">定义在其上显示外接程序命令的功能区选项卡。这可以是默认的选项卡（“**主页**”、“**消息**”或“**会议**”），或是由外接程序定义的自定义选项卡。此元素是必需的。</span><span class="sxs-lookup"><span data-stu-id="13bc9-p101">Defines the ribbon tab on which your add-in command appears. This can either be the default tab (either  **Home**,  **Message**, or  **Meeting**), or a custom tab defined by the add-in. This element is required.</span></span>

## <a name="child-elements"></a><span data-ttu-id="13bc9-105">子元素</span><span class="sxs-lookup"><span data-stu-id="13bc9-105">Child elements</span></span>

|  <span data-ttu-id="13bc9-106">元素</span><span class="sxs-lookup"><span data-stu-id="13bc9-106">Element</span></span> |  <span data-ttu-id="13bc9-107">必需</span><span class="sxs-lookup"><span data-stu-id="13bc9-107">Required</span></span>  |  <span data-ttu-id="13bc9-108">说明</span><span class="sxs-lookup"><span data-stu-id="13bc9-108">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="13bc9-109">Group</span><span class="sxs-lookup"><span data-stu-id="13bc9-109">Group</span></span>      | <span data-ttu-id="13bc9-110">是</span><span class="sxs-lookup"><span data-stu-id="13bc9-110">Yes</span></span> |  <span data-ttu-id="13bc9-p102">定义一组命令。对于每个外接程序，只能将一个组添加到默认选项卡。</span><span class="sxs-lookup"><span data-stu-id="13bc9-p102">Defines a group of commands. You can add only one group per add-in to the default tab.</span></span>  |

<span data-ttu-id="13bc9-p103">下面是主机的有效选项卡 `id` 值。以**粗体** 显示的值在桌面和联机状态中均受支持（例如，适用于 Windows 的 Word 2016 和 Word Online）。</span><span class="sxs-lookup"><span data-stu-id="13bc9-p103">The following are valid tab `id` values by host. Values in **bold** are supported in both desktop and online (for example, Word 2016 for Windows and Word Online).</span></span> 

### <a name="outlook"></a><span data-ttu-id="13bc9-115">Outlook</span><span class="sxs-lookup"><span data-stu-id="13bc9-115">Outlook</span></span> 

- <span data-ttu-id="13bc9-116">**TabDefault**</span><span class="sxs-lookup"><span data-stu-id="13bc9-116">**TabDefault**</span></span>

### <a name="word"></a><span data-ttu-id="13bc9-117">Word</span><span class="sxs-lookup"><span data-stu-id="13bc9-117">Word</span></span>

- <span data-ttu-id="13bc9-118">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="13bc9-118">**TabHome**</span></span>
- <span data-ttu-id="13bc9-119">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="13bc9-119">**TabInsert**</span></span>
- <span data-ttu-id="13bc9-120">TabWordDesign</span><span class="sxs-lookup"><span data-stu-id="13bc9-120">TabWordDesign</span></span>
- <span data-ttu-id="13bc9-121">**TabPageLayoutWord**</span><span class="sxs-lookup"><span data-stu-id="13bc9-121">**TabPageLayoutWord**</span></span>
- <span data-ttu-id="13bc9-122">TabReferences</span><span class="sxs-lookup"><span data-stu-id="13bc9-122">TabReferences</span></span>
- <span data-ttu-id="13bc9-123">TabMailings</span><span class="sxs-lookup"><span data-stu-id="13bc9-123">TabMailings</span></span>
- <span data-ttu-id="13bc9-124">TabReviewWord</span><span class="sxs-lookup"><span data-stu-id="13bc9-124">TabReviewWord</span></span>
- <span data-ttu-id="13bc9-125">**TabView**</span><span class="sxs-lookup"><span data-stu-id="13bc9-125">**TabView**</span></span>
- <span data-ttu-id="13bc9-126">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="13bc9-126">TabDeveloper</span></span>
- <span data-ttu-id="13bc9-127">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="13bc9-127">TabAddIns</span></span>
- <span data-ttu-id="13bc9-128">TabBlogPost</span><span class="sxs-lookup"><span data-stu-id="13bc9-128">TabBlogPost</span></span>
- <span data-ttu-id="13bc9-129">TabBlogInsert</span><span class="sxs-lookup"><span data-stu-id="13bc9-129">TabBlogInsert</span></span>
- <span data-ttu-id="13bc9-130">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="13bc9-130">TabPrintPreview</span></span>
- <span data-ttu-id="13bc9-131">TabOutlining</span><span class="sxs-lookup"><span data-stu-id="13bc9-131">TabOutlining</span></span>
- <span data-ttu-id="13bc9-132">TabConflicts</span><span class="sxs-lookup"><span data-stu-id="13bc9-132">TabConflicts</span></span>
- <span data-ttu-id="13bc9-133">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="13bc9-133">TabBackgroundRemoval</span></span>
- <span data-ttu-id="13bc9-134">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="13bc9-134">TabBroadcastPresentation</span></span>

### <a name="excel"></a><span data-ttu-id="13bc9-135">Excel</span><span class="sxs-lookup"><span data-stu-id="13bc9-135">Excel</span></span>

- <span data-ttu-id="13bc9-136">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="13bc9-136">**TabHome**</span></span>
- <span data-ttu-id="13bc9-137">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="13bc9-137">**TabInsert**</span></span>
- <span data-ttu-id="13bc9-138">TabPageLayoutExcel</span><span class="sxs-lookup"><span data-stu-id="13bc9-138">TabPageLayoutExcel</span></span>
- <span data-ttu-id="13bc9-139">TabFormulas</span><span class="sxs-lookup"><span data-stu-id="13bc9-139">TabFormulas</span></span>
- <span data-ttu-id="13bc9-140">**TabData**</span><span class="sxs-lookup"><span data-stu-id="13bc9-140">**TabData**</span></span>
- <span data-ttu-id="13bc9-141">**TabReview**</span><span class="sxs-lookup"><span data-stu-id="13bc9-141">**TabReview**</span></span>
- <span data-ttu-id="13bc9-142">**TabView**</span><span class="sxs-lookup"><span data-stu-id="13bc9-142">**TabView**</span></span>
- <span data-ttu-id="13bc9-143">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="13bc9-143">TabDeveloper</span></span>
- <span data-ttu-id="13bc9-144">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="13bc9-144">TabAddIns</span></span>
- <span data-ttu-id="13bc9-145">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="13bc9-145">TabPrintPreview</span></span>
- <span data-ttu-id="13bc9-146">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="13bc9-146">TabBackgroundRemoval</span></span> 

### <a name="powerpoint"></a><span data-ttu-id="13bc9-147">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="13bc9-147">PowerPoint</span></span>

- <span data-ttu-id="13bc9-148">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="13bc9-148">**TabHome**</span></span>
- <span data-ttu-id="13bc9-149">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="13bc9-149">**TabInsert**</span></span>
- <span data-ttu-id="13bc9-150">**TabDesign**</span><span class="sxs-lookup"><span data-stu-id="13bc9-150">**TabDesign**</span></span>
- <span data-ttu-id="13bc9-151">**TabTransitions**</span><span class="sxs-lookup"><span data-stu-id="13bc9-151">**TabTransitions**</span></span>
- <span data-ttu-id="13bc9-152">**TabAnimations**</span><span class="sxs-lookup"><span data-stu-id="13bc9-152">**TabAnimations**</span></span>
- <span data-ttu-id="13bc9-153">TabSlideShow</span><span class="sxs-lookup"><span data-stu-id="13bc9-153">TabSlideShow</span></span>
- <span data-ttu-id="13bc9-154">TabReview</span><span class="sxs-lookup"><span data-stu-id="13bc9-154">TabReview</span></span>
- <span data-ttu-id="13bc9-155">**TabView**</span><span class="sxs-lookup"><span data-stu-id="13bc9-155">**TabView**</span></span>
- <span data-ttu-id="13bc9-156">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="13bc9-156">TabDeveloper</span></span>
- <span data-ttu-id="13bc9-157">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="13bc9-157">TabAddIns</span></span>
- <span data-ttu-id="13bc9-158">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="13bc9-158">TabPrintPreview</span></span>
- <span data-ttu-id="13bc9-159">TabMerge</span><span class="sxs-lookup"><span data-stu-id="13bc9-159">TabMerge</span></span>
- <span data-ttu-id="13bc9-160">TabGrayscale</span><span class="sxs-lookup"><span data-stu-id="13bc9-160">TabGrayscale</span></span>
- <span data-ttu-id="13bc9-161">TabBlackAndWhite</span><span class="sxs-lookup"><span data-stu-id="13bc9-161">TabBlackAndWhite</span></span>
- <span data-ttu-id="13bc9-162">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="13bc9-162">TabBroadcastPresentation</span></span>
- <span data-ttu-id="13bc9-163">TabSlideMaster</span><span class="sxs-lookup"><span data-stu-id="13bc9-163">TabSlideMaster</span></span>
- <span data-ttu-id="13bc9-164">TabHandoutMaster</span><span class="sxs-lookup"><span data-stu-id="13bc9-164">TabHandoutMaster</span></span>
- <span data-ttu-id="13bc9-165">TabNotesMaster</span><span class="sxs-lookup"><span data-stu-id="13bc9-165">TabNotesMaster</span></span>
- <span data-ttu-id="13bc9-166">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="13bc9-166">TabBackgroundRemoval</span></span>
- <span data-ttu-id="13bc9-167">TabSlideMasterHome</span><span class="sxs-lookup"><span data-stu-id="13bc9-167">TabSlideMasterHome</span></span>

### <a name="onenote"></a><span data-ttu-id="13bc9-168">OneNote</span><span class="sxs-lookup"><span data-stu-id="13bc9-168">OneNote</span></span>

- <span data-ttu-id="13bc9-169">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="13bc9-169">**TabHome**</span></span>
- <span data-ttu-id="13bc9-170">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="13bc9-170">**TabInsert**</span></span>
- <span data-ttu-id="13bc9-171">**TabView**</span><span class="sxs-lookup"><span data-stu-id="13bc9-171">**TabView**</span></span>
- <span data-ttu-id="13bc9-172">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="13bc9-172">TabDeveloper</span></span>
- <span data-ttu-id="13bc9-173">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="13bc9-173">TabAddIns</span></span>

## <a name="group"></a><span data-ttu-id="13bc9-174">组</span><span class="sxs-lookup"><span data-stu-id="13bc9-174">Group</span></span>

<span data-ttu-id="13bc9-p104">选项卡中的一组 UI 扩展点。一组可以有多达六个控件。需要 **id** 属性且每个 **id** 在清单内必须是唯一的。**id** 是一个最大长度为 125 个字符的字符串。查看 [Group 元素](group.md)。</span><span class="sxs-lookup"><span data-stu-id="13bc9-p104">A group of UI extension points in a tab. A group can have up to six controls. The  **id** attribute is required and each **id** must be unique within the manifest. The **id** is a string with a maximum of 125 characters. See [Group element](group.md).</span></span>

## <a name="officetab-example"></a><span data-ttu-id="13bc9-179">OfficeTab 示例</span><span class="sxs-lookup"><span data-stu-id="13bc9-179">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
    <Group id="msgreadTabMessage.grp1">
        <!-- Group Definition -->
    </Group>
  </OfficeTab>
</ExtensionPoint>
```
