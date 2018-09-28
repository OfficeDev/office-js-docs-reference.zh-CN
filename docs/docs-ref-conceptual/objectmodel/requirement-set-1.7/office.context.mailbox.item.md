
# <a name="item"></a><span data-ttu-id="3dace-101">item</span><span class="sxs-lookup"><span data-stu-id="3dace-101">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="3dace-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="3dace-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="3dace-p101">`item` 命名空间用于访问当前选定的邮件、会议请求或约会。可以通过使用 [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook17officemailboxenumsitemtype) 属性确定 `item` 的类型。</span><span class="sxs-lookup"><span data-stu-id="3dace-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook17officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="3dace-105">Requirements</span><span class="sxs-lookup"><span data-stu-id="3dace-105">Requirements</span></span>

|<span data-ttu-id="3dace-106">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-106">Requirement</span></span>|<span data-ttu-id="3dace-107">值</span><span class="sxs-lookup"><span data-stu-id="3dace-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="3dace-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3dace-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="3dace-109">1.0</span><span class="sxs-lookup"><span data-stu-id="3dace-109">1.0</span></span>|
|[<span data-ttu-id="3dace-110">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3dace-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="3dace-111">受限</span><span class="sxs-lookup"><span data-stu-id="3dace-111">Restricted</span></span>|
|[<span data-ttu-id="3dace-112">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3dace-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="3dace-113">Compose 或 Read</span><span class="sxs-lookup"><span data-stu-id="3dace-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="3dace-114">成员和方法</span><span class="sxs-lookup"><span data-stu-id="3dace-114">Members and methods</span></span>

| <span data-ttu-id="3dace-115">成员</span><span class="sxs-lookup"><span data-stu-id="3dace-115">Member</span></span> | <span data-ttu-id="3dace-116">类型</span><span class="sxs-lookup"><span data-stu-id="3dace-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="3dace-117">attachments</span><span class="sxs-lookup"><span data-stu-id="3dace-117">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlook17officeattachmentdetails) | <span data-ttu-id="3dace-118">成员</span><span class="sxs-lookup"><span data-stu-id="3dace-118">Member</span></span> |
| [<span data-ttu-id="3dace-119">bcc</span><span class="sxs-lookup"><span data-stu-id="3dace-119">bcc</span></span>](#bcc-recipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="3dace-120">成员</span><span class="sxs-lookup"><span data-stu-id="3dace-120">Member</span></span> |
| [<span data-ttu-id="3dace-121">body</span><span class="sxs-lookup"><span data-stu-id="3dace-121">body</span></span>](#body-bodyjavascriptapioutlook17officebody) | <span data-ttu-id="3dace-122">成员</span><span class="sxs-lookup"><span data-stu-id="3dace-122">Member</span></span> |
| [<span data-ttu-id="3dace-123">cc</span><span class="sxs-lookup"><span data-stu-id="3dace-123">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="3dace-124">成员</span><span class="sxs-lookup"><span data-stu-id="3dace-124">Member</span></span> |
| [<span data-ttu-id="3dace-125">conversationId</span><span class="sxs-lookup"><span data-stu-id="3dace-125">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="3dace-126">成员</span><span class="sxs-lookup"><span data-stu-id="3dace-126">Member</span></span> |
| [<span data-ttu-id="3dace-127">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="3dace-127">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="3dace-128">成员</span><span class="sxs-lookup"><span data-stu-id="3dace-128">Member</span></span> |
| [<span data-ttu-id="3dace-129">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="3dace-129">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="3dace-130">成员</span><span class="sxs-lookup"><span data-stu-id="3dace-130">Member</span></span> |
| [<span data-ttu-id="3dace-131">end</span><span class="sxs-lookup"><span data-stu-id="3dace-131">end</span></span>](#end-datetimejavascriptapioutlook17officetime) | <span data-ttu-id="3dace-132">成员</span><span class="sxs-lookup"><span data-stu-id="3dace-132">Member</span></span> |
| [<span data-ttu-id="3dace-133">from</span><span class="sxs-lookup"><span data-stu-id="3dace-133">from</span></span>](#from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom) | <span data-ttu-id="3dace-134">成员</span><span class="sxs-lookup"><span data-stu-id="3dace-134">Member</span></span> |
| [<span data-ttu-id="3dace-135">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="3dace-135">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="3dace-136">成员</span><span class="sxs-lookup"><span data-stu-id="3dace-136">Member</span></span> |
| [<span data-ttu-id="3dace-137">itemClass</span><span class="sxs-lookup"><span data-stu-id="3dace-137">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="3dace-138">成员</span><span class="sxs-lookup"><span data-stu-id="3dace-138">Member</span></span> |
| [<span data-ttu-id="3dace-139">itemId</span><span class="sxs-lookup"><span data-stu-id="3dace-139">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="3dace-140">成员</span><span class="sxs-lookup"><span data-stu-id="3dace-140">Member</span></span> |
| [<span data-ttu-id="3dace-141">itemType</span><span class="sxs-lookup"><span data-stu-id="3dace-141">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlook17officemailboxenumsitemtype) | <span data-ttu-id="3dace-142">成员</span><span class="sxs-lookup"><span data-stu-id="3dace-142">Member</span></span> |
| [<span data-ttu-id="3dace-143">location</span><span class="sxs-lookup"><span data-stu-id="3dace-143">location</span></span>](#location-stringlocationjavascriptapioutlook17officelocation) | <span data-ttu-id="3dace-144">成员</span><span class="sxs-lookup"><span data-stu-id="3dace-144">Member</span></span> |
| [<span data-ttu-id="3dace-145">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="3dace-145">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="3dace-146">成员</span><span class="sxs-lookup"><span data-stu-id="3dace-146">Member</span></span> |
| [<span data-ttu-id="3dace-147">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="3dace-147">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlook17officenotificationmessages) | <span data-ttu-id="3dace-148">成员</span><span class="sxs-lookup"><span data-stu-id="3dace-148">Member</span></span> |
| [<span data-ttu-id="3dace-149">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="3dace-149">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="3dace-150">成员</span><span class="sxs-lookup"><span data-stu-id="3dace-150">Member</span></span> |
| [<span data-ttu-id="3dace-151">organizer</span><span class="sxs-lookup"><span data-stu-id="3dace-151">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsorganizerjavascriptapioutlook17officeorganizer) | <span data-ttu-id="3dace-152">成员</span><span class="sxs-lookup"><span data-stu-id="3dace-152">Member</span></span> |
| [<span data-ttu-id="3dace-153">recurrence</span><span class="sxs-lookup"><span data-stu-id="3dace-153">recurrence</span></span>](#nullable-recurrence-recurrencejavascriptapioutlook17officerecurrence) | <span data-ttu-id="3dace-154">成员</span><span class="sxs-lookup"><span data-stu-id="3dace-154">Member</span></span> |
| [<span data-ttu-id="3dace-155">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="3dace-155">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="3dace-156">成员</span><span class="sxs-lookup"><span data-stu-id="3dace-156">Member</span></span> |
| [<span data-ttu-id="3dace-157">sender</span><span class="sxs-lookup"><span data-stu-id="3dace-157">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetails) | <span data-ttu-id="3dace-158">成员</span><span class="sxs-lookup"><span data-stu-id="3dace-158">Member</span></span> |
| [<span data-ttu-id="3dace-159">可能指向</span><span class="sxs-lookup"><span data-stu-id="3dace-159">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="3dace-160">成员</span><span class="sxs-lookup"><span data-stu-id="3dace-160">Member</span></span> |
| [<span data-ttu-id="3dace-161">start</span><span class="sxs-lookup"><span data-stu-id="3dace-161">start</span></span>](#start-datetimejavascriptapioutlook17officetime) | <span data-ttu-id="3dace-162">成员</span><span class="sxs-lookup"><span data-stu-id="3dace-162">Member</span></span> |
| [<span data-ttu-id="3dace-163">subject</span><span class="sxs-lookup"><span data-stu-id="3dace-163">subject</span></span>](#subject-stringsubjectjavascriptapioutlook17officesubject) | <span data-ttu-id="3dace-164">成员</span><span class="sxs-lookup"><span data-stu-id="3dace-164">Member</span></span> |
| [<span data-ttu-id="3dace-165">to</span><span class="sxs-lookup"><span data-stu-id="3dace-165">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="3dace-166">成员</span><span class="sxs-lookup"><span data-stu-id="3dace-166">Member</span></span> |
| [<span data-ttu-id="3dace-167">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="3dace-167">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="3dace-168">方法</span><span class="sxs-lookup"><span data-stu-id="3dace-168">Method</span></span> |
| [<span data-ttu-id="3dace-169">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="3dace-169">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="3dace-170">方法</span><span class="sxs-lookup"><span data-stu-id="3dace-170">Method</span></span> |
| [<span data-ttu-id="3dace-171">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="3dace-171">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="3dace-172">方法</span><span class="sxs-lookup"><span data-stu-id="3dace-172">Method</span></span> |
| [<span data-ttu-id="3dace-173">close</span><span class="sxs-lookup"><span data-stu-id="3dace-173">close</span></span>](#close) | <span data-ttu-id="3dace-174">方法</span><span class="sxs-lookup"><span data-stu-id="3dace-174">Method</span></span> |
| [<span data-ttu-id="3dace-175">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="3dace-175">displayReplyAllForm</span></span>](#displayreplyallformformdata) | <span data-ttu-id="3dace-176">方法</span><span class="sxs-lookup"><span data-stu-id="3dace-176">Method</span></span> |
| [<span data-ttu-id="3dace-177">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="3dace-177">displayReplyForm</span></span>](#displayreplyformformdata) | <span data-ttu-id="3dace-178">方法</span><span class="sxs-lookup"><span data-stu-id="3dace-178">Method</span></span> |
| [<span data-ttu-id="3dace-179">getEntities</span><span class="sxs-lookup"><span data-stu-id="3dace-179">getEntities</span></span>](#getentities--entitiesjavascriptapioutlook17officeentities) | <span data-ttu-id="3dace-180">方法</span><span class="sxs-lookup"><span data-stu-id="3dace-180">Method</span></span> |
| [<span data-ttu-id="3dace-181">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="3dace-181">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion) | <span data-ttu-id="3dace-182">方法</span><span class="sxs-lookup"><span data-stu-id="3dace-182">Method</span></span> |
| [<span data-ttu-id="3dace-183">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="3dace-183">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion) | <span data-ttu-id="3dace-184">方法</span><span class="sxs-lookup"><span data-stu-id="3dace-184">Method</span></span> |
| [<span data-ttu-id="3dace-185">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="3dace-185">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="3dace-186">方法</span><span class="sxs-lookup"><span data-stu-id="3dace-186">Method</span></span> |
| [<span data-ttu-id="3dace-187">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="3dace-187">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="3dace-188">方法</span><span class="sxs-lookup"><span data-stu-id="3dace-188">Method</span></span> |
| [<span data-ttu-id="3dace-189">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="3dace-189">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="3dace-190">方法</span><span class="sxs-lookup"><span data-stu-id="3dace-190">Method</span></span> |
| [<span data-ttu-id="3dace-191">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="3dace-191">getSelectedEntities</span></span>](#getselectedentities--entitiesjavascriptapioutlook17officeentities) | <span data-ttu-id="3dace-192">方法</span><span class="sxs-lookup"><span data-stu-id="3dace-192">Method</span></span> |
| [<span data-ttu-id="3dace-193">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="3dace-193">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="3dace-194">方法</span><span class="sxs-lookup"><span data-stu-id="3dace-194">Method</span></span> |
| [<span data-ttu-id="3dace-195">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="3dace-195">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="3dace-196">方法</span><span class="sxs-lookup"><span data-stu-id="3dace-196">Method</span></span> |
| [<span data-ttu-id="3dace-197">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="3dace-197">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="3dace-198">方法</span><span class="sxs-lookup"><span data-stu-id="3dace-198">Method</span></span> |
| [<span data-ttu-id="3dace-199">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="3dace-199">removeHandlerAsync</span></span>](#removehandlerasynceventtype-handler-options-callback) | <span data-ttu-id="3dace-200">方法</span><span class="sxs-lookup"><span data-stu-id="3dace-200">Method</span></span> |
| [<span data-ttu-id="3dace-201">saveAsync</span><span class="sxs-lookup"><span data-stu-id="3dace-201">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="3dace-202">方法</span><span class="sxs-lookup"><span data-stu-id="3dace-202">Method</span></span> |
| [<span data-ttu-id="3dace-203">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="3dace-203">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="3dace-204">方法</span><span class="sxs-lookup"><span data-stu-id="3dace-204">Method</span></span> |

### <a name="example"></a><span data-ttu-id="3dace-205">示例</span><span class="sxs-lookup"><span data-stu-id="3dace-205">Example</span></span>

<span data-ttu-id="3dace-206">以下 JavaScript 代码示例显示了如何访问 Outlook 中当前项目的 `subject` 属性。</span><span class="sxs-lookup"><span data-stu-id="3dace-206">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```
// The initialize function is required for all apps.
Office.initialize = function () {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    var item = Office.context.mailbox.item;
    var subject = item.subject;
    // Continue with processing the subject of the current item,
    // which can be a message or appointment.
    });
}
```

### <a name="members"></a><span data-ttu-id="3dace-207">成员</span><span class="sxs-lookup"><span data-stu-id="3dace-207">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook17officeattachmentdetails"></a><span data-ttu-id="3dace-208">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="3dace-208">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span></span>

<span data-ttu-id="3dace-p102">获取项目的附件数组。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="3dace-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="3dace-211">某些类型的文件被阻止 Outlook 由于潜在的安全问题，并因此不返回。</span><span class="sxs-lookup"><span data-stu-id="3dace-211">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="3dace-212">有关详细信息，请参阅[在 Outlook 中被阻止的附件](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)。</span><span class="sxs-lookup"><span data-stu-id="3dace-212">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="3dace-213">类型:</span><span class="sxs-lookup"><span data-stu-id="3dace-213">Type:</span></span>

*   <span data-ttu-id="3dace-214">Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="3dace-214">Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="3dace-215">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-215">Requirements</span></span>

|<span data-ttu-id="3dace-216">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-216">Requirement</span></span>|<span data-ttu-id="3dace-217">值</span><span class="sxs-lookup"><span data-stu-id="3dace-217">Value</span></span>|
|---|---|
|[<span data-ttu-id="3dace-218">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3dace-218">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="3dace-219">1.0</span><span class="sxs-lookup"><span data-stu-id="3dace-219">1.0</span></span>|
|[<span data-ttu-id="3dace-220">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3dace-220">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="3dace-221">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3dace-221">ReadItem</span></span>|
|[<span data-ttu-id="3dace-222">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3dace-222">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="3dace-223">阅读</span><span class="sxs-lookup"><span data-stu-id="3dace-223">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3dace-224">示例</span><span class="sxs-lookup"><span data-stu-id="3dace-224">Example</span></span>

<span data-ttu-id="3dace-225">以下代码使用当前项目上所有附件的详细信息构成 HTML 字符串。</span><span class="sxs-lookup"><span data-stu-id="3dace-225">The following code builds an HTML string with details of all attachments on the current item.</span></span>

```
var _Item = Office.context.mailbox.item;
var outputString = "";

if (_Item.attachments.length > 0) {
  for (i = 0 ; i < _Item.attachments.length ; i++) {
    var _att = _Item.attachments[i];
    outputString += "<BR>" + i + ". Name: ";
    outputString += _att.name;
    outputString += "<BR>ID: " + _att.id;
    outputString += "<BR>contentType: " + _att.contentType;
    outputString += "<BR>size: " + _att.size;
    outputString += "<BR>attachmentType: " + _att.attachmentType;
    outputString += "<BR>isInline: " + _att.isInline;
  }
}

// Do something with outputString
```

####  <a name="bcc-recipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="3dace-226">bcc :[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="3dace-226">bcc :[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="3dace-227">获取一个对象，提供用于获取或更新邮件的密件抄送 （密件抄送副本） 行上的收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="3dace-227">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="3dace-228">仅限撰写模式。</span><span class="sxs-lookup"><span data-stu-id="3dace-228">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="3dace-229">类型：</span><span class="sxs-lookup"><span data-stu-id="3dace-229">Type:</span></span>

*   [<span data-ttu-id="3dace-230">收件人</span><span class="sxs-lookup"><span data-stu-id="3dace-230">Recipients</span></span>](/javascript/api/outlook_1_7/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="3dace-231">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-231">Requirements</span></span>

|<span data-ttu-id="3dace-232">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-232">Requirement</span></span>|<span data-ttu-id="3dace-233">值</span><span class="sxs-lookup"><span data-stu-id="3dace-233">Value</span></span>|
|---|---|
|[<span data-ttu-id="3dace-234">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3dace-234">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="3dace-235">1.1</span><span class="sxs-lookup"><span data-stu-id="3dace-235">1.1</span></span>|
|[<span data-ttu-id="3dace-236">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3dace-236">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="3dace-237">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3dace-237">ReadItem</span></span>|
|[<span data-ttu-id="3dace-238">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3dace-238">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="3dace-239">撰写</span><span class="sxs-lookup"><span data-stu-id="3dace-239">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="3dace-240">示例</span><span class="sxs-lookup"><span data-stu-id="3dace-240">Example</span></span>

```
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook17officebody"></a><span data-ttu-id="3dace-241">body :[Body](/javascript/api/outlook_1_7/office.body)</span><span class="sxs-lookup"><span data-stu-id="3dace-241">body :[Body](/javascript/api/outlook_1_7/office.body)</span></span>

<span data-ttu-id="3dace-242">获取一个提供用于处理项目正文的方法的对象。</span><span class="sxs-lookup"><span data-stu-id="3dace-242">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="3dace-243">类型：</span><span class="sxs-lookup"><span data-stu-id="3dace-243">Type:</span></span>

*   [<span data-ttu-id="3dace-244">Body</span><span class="sxs-lookup"><span data-stu-id="3dace-244">Body</span></span>](/javascript/api/outlook_1_7/office.body)

##### <a name="requirements"></a><span data-ttu-id="3dace-245">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-245">Requirements</span></span>

|<span data-ttu-id="3dace-246">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-246">Requirement</span></span>|<span data-ttu-id="3dace-247">值</span><span class="sxs-lookup"><span data-stu-id="3dace-247">Value</span></span>|
|---|---|
|[<span data-ttu-id="3dace-248">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3dace-248">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="3dace-249">1.1</span><span class="sxs-lookup"><span data-stu-id="3dace-249">1.1</span></span>|
|[<span data-ttu-id="3dace-250">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3dace-250">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="3dace-251">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3dace-251">ReadItem</span></span>|
|[<span data-ttu-id="3dace-252">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3dace-252">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="3dace-253">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3dace-253">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="3dace-254">抄送： 数组。 <[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[收件人](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="3dace-254">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="3dace-255">提供对邮件的抄送 (cc) 收件人访问。</span><span class="sxs-lookup"><span data-stu-id="3dace-255">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="3dace-256">对象的类型和访问级别取决于当前项的模式。</span><span class="sxs-lookup"><span data-stu-id="3dace-256">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="3dace-257">阅读模式</span><span class="sxs-lookup"><span data-stu-id="3dace-257">Read mode</span></span>

<span data-ttu-id="3dace-p106">`cc` 属性返回包含邮件的**抄送**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="3dace-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="3dace-260">撰写模式</span><span class="sxs-lookup"><span data-stu-id="3dace-260">Compose mode</span></span>

<span data-ttu-id="3dace-261">`cc`属性返回`Recipients`提供方法用于获取或更新的邮件**Cc**行上的收件人对象。</span><span class="sxs-lookup"><span data-stu-id="3dace-261">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="3dace-262">类型：</span><span class="sxs-lookup"><span data-stu-id="3dace-262">Type:</span></span>

*   <span data-ttu-id="3dace-263">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="3dace-263">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="3dace-264">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-264">Requirements</span></span>

|<span data-ttu-id="3dace-265">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-265">Requirement</span></span>|<span data-ttu-id="3dace-266">值</span><span class="sxs-lookup"><span data-stu-id="3dace-266">Value</span></span>|
|---|---|
|[<span data-ttu-id="3dace-267">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3dace-267">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="3dace-268">1.0</span><span class="sxs-lookup"><span data-stu-id="3dace-268">1.0</span></span>|
|[<span data-ttu-id="3dace-269">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3dace-269">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="3dace-270">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3dace-270">ReadItem</span></span>|
|[<span data-ttu-id="3dace-271">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3dace-271">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="3dace-272">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3dace-272">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="3dace-273">示例</span><span class="sxs-lookup"><span data-stu-id="3dace-273">Example</span></span>

```
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="3dace-274">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="3dace-274">(nullable) conversationId :String</span></span>

<span data-ttu-id="3dace-275">获取包含特定消息的电子邮件会话的标识符。</span><span class="sxs-lookup"><span data-stu-id="3dace-275">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="3dace-p107">如果在阅读窗体或撰写窗体的回复中激活邮件应用程序，则此属性可以获得一个整数值。如果用户随后更改了回复邮件的主题（若发送回复），则该邮件的对话 ID 将改变且之前获取的值将不适用。</span><span class="sxs-lookup"><span data-stu-id="3dace-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="3dace-p108">对于撰写窗体的新项目，此属性获得一个 null 值。如果用户设置一个主题并保存该项目，`conversationId` 属性将返回一个值。</span><span class="sxs-lookup"><span data-stu-id="3dace-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="3dace-280">类型:</span><span class="sxs-lookup"><span data-stu-id="3dace-280">Type:</span></span>

*   <span data-ttu-id="3dace-281">String</span><span class="sxs-lookup"><span data-stu-id="3dace-281">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="3dace-282">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-282">Requirements</span></span>

|<span data-ttu-id="3dace-283">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-283">Requirement</span></span>|<span data-ttu-id="3dace-284">值</span><span class="sxs-lookup"><span data-stu-id="3dace-284">Value</span></span>|
|---|---|
|[<span data-ttu-id="3dace-285">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3dace-285">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="3dace-286">1.0</span><span class="sxs-lookup"><span data-stu-id="3dace-286">1.0</span></span>|
|[<span data-ttu-id="3dace-287">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3dace-287">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="3dace-288">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3dace-288">ReadItem</span></span>|
|[<span data-ttu-id="3dace-289">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3dace-289">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="3dace-290">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3dace-290">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="3dace-291">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="3dace-291">dateTimeCreated :Date</span></span>

<span data-ttu-id="3dace-p109">获取项目创建的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="3dace-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="3dace-294">类型：</span><span class="sxs-lookup"><span data-stu-id="3dace-294">Type:</span></span>

*   <span data-ttu-id="3dace-295">日期</span><span class="sxs-lookup"><span data-stu-id="3dace-295">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="3dace-296">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-296">Requirements</span></span>

|<span data-ttu-id="3dace-297">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-297">Requirement</span></span>|<span data-ttu-id="3dace-298">值</span><span class="sxs-lookup"><span data-stu-id="3dace-298">Value</span></span>|
|---|---|
|[<span data-ttu-id="3dace-299">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3dace-299">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="3dace-300">1.0</span><span class="sxs-lookup"><span data-stu-id="3dace-300">1.0</span></span>|
|[<span data-ttu-id="3dace-301">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3dace-301">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="3dace-302">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3dace-302">ReadItem</span></span>|
|[<span data-ttu-id="3dace-303">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3dace-303">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="3dace-304">阅读</span><span class="sxs-lookup"><span data-stu-id="3dace-304">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3dace-305">示例</span><span class="sxs-lookup"><span data-stu-id="3dace-305">Example</span></span>

```
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="3dace-306">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="3dace-306">dateTimeModified :Date</span></span>

<span data-ttu-id="3dace-p110">获取项目最近一次修改的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="3dace-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="3dace-309">此成员不支持适用于 iOS 的 Outlook 或 Outlook for Android。</span><span class="sxs-lookup"><span data-stu-id="3dace-309">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="3dace-310">类型:</span><span class="sxs-lookup"><span data-stu-id="3dace-310">Type:</span></span>

*   <span data-ttu-id="3dace-311">日期</span><span class="sxs-lookup"><span data-stu-id="3dace-311">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="3dace-312">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-312">Requirements</span></span>

|<span data-ttu-id="3dace-313">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-313">Requirement</span></span>|<span data-ttu-id="3dace-314">值</span><span class="sxs-lookup"><span data-stu-id="3dace-314">Value</span></span>|
|---|---|
|[<span data-ttu-id="3dace-315">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3dace-315">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="3dace-316">1.0</span><span class="sxs-lookup"><span data-stu-id="3dace-316">1.0</span></span>|
|[<span data-ttu-id="3dace-317">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3dace-317">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="3dace-318">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3dace-318">ReadItem</span></span>|
|[<span data-ttu-id="3dace-319">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3dace-319">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="3dace-320">阅读</span><span class="sxs-lookup"><span data-stu-id="3dace-320">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3dace-321">示例</span><span class="sxs-lookup"><span data-stu-id="3dace-321">Example</span></span>

```
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook17officetime"></a><span data-ttu-id="3dace-322">end :Date|[Time](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="3dace-322">end :Date|[Time](/javascript/api/outlook_1_7/office.time)</span></span>

<span data-ttu-id="3dace-323">获取或设置约会结束的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="3dace-323">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="3dace-p111">将 `end` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime) 方法将 end 属性值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="3dace-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="3dace-326">阅读模式</span><span class="sxs-lookup"><span data-stu-id="3dace-326">Read mode</span></span>

<span data-ttu-id="3dace-327">`end` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="3dace-327">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="3dace-328">撰写模式</span><span class="sxs-lookup"><span data-stu-id="3dace-328">Compose mode</span></span>

<span data-ttu-id="3dace-329">`end` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="3dace-329">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="3dace-330">使用 [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) 方法设置结束时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="3dace-330">When you use the [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="3dace-331">类型：</span><span class="sxs-lookup"><span data-stu-id="3dace-331">Type:</span></span>

*   <span data-ttu-id="3dace-332">Date | [Time](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="3dace-332">Date | [Time](/javascript/api/outlook_1_7/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="3dace-333">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-333">Requirements</span></span>

|<span data-ttu-id="3dace-334">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-334">Requirement</span></span>|<span data-ttu-id="3dace-335">值</span><span class="sxs-lookup"><span data-stu-id="3dace-335">Value</span></span>|
|---|---|
|[<span data-ttu-id="3dace-336">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3dace-336">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="3dace-337">1.0</span><span class="sxs-lookup"><span data-stu-id="3dace-337">1.0</span></span>|
|[<span data-ttu-id="3dace-338">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3dace-338">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="3dace-339">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3dace-339">ReadItem</span></span>|
|[<span data-ttu-id="3dace-340">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3dace-340">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="3dace-341">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3dace-341">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="3dace-342">示例</span><span class="sxs-lookup"><span data-stu-id="3dace-342">Example</span></span>

<span data-ttu-id="3dace-343">以下示例通过使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) 方法，设置撰写模式下约会的结束时间。</span><span class="sxs-lookup"><span data-stu-id="3dace-343">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```
var endTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used
  // in the callback
     asyncContext: {verb:"Set"}
}
Office.context.mailbox.item.end.setAsync(endTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function
    console.debug("End Time " + result.asyncContext.verb);
  }
});
```

#### <a name="from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom"></a><span data-ttu-id="3dace-344">从：[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[从](/javascript/api/outlook_1_7/office.from)</span><span class="sxs-lookup"><span data-stu-id="3dace-344">from :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[From](/javascript/api/outlook_1_7/office.from)</span></span>

<span data-ttu-id="3dace-345">获取邮件发件人的电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="3dace-345">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="3dace-p112">`from` 和 [`sender`](#sender-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetails) 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="3dace-p112">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="3dace-348">`recipientType`属性`EmailAddressDetails`对象在`from`属性是`undefined`。</span><span class="sxs-lookup"><span data-stu-id="3dace-348">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="3dace-349">阅读模式</span><span class="sxs-lookup"><span data-stu-id="3dace-349">Read mode</span></span>

<span data-ttu-id="3dace-350">`from`属性返回`EmailAddressDetails`对象。</span><span class="sxs-lookup"><span data-stu-id="3dace-350">The `from` property returns an `EmailAddressDetails` object.</span></span>

```
var subject = Office.context.mailbox.item.from;
```

##### <a name="compose-mode"></a><span data-ttu-id="3dace-351">撰写模式</span><span class="sxs-lookup"><span data-stu-id="3dace-351">Compose mode</span></span>

<span data-ttu-id="3dace-352">`from`属性返回`From`对象，它提供一个方法来获取值。</span><span class="sxs-lookup"><span data-stu-id="3dace-352">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="3dace-353">类型:</span><span class="sxs-lookup"><span data-stu-id="3dace-353">Type:</span></span>

*   <span data-ttu-id="3dace-354">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [从](/javascript/api/outlook_1_7/office.from)</span><span class="sxs-lookup"><span data-stu-id="3dace-354">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [From](/javascript/api/outlook_1_7/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="3dace-355">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-355">Requirements</span></span>

|<span data-ttu-id="3dace-356">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-356">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="3dace-357">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3dace-357">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="3dace-358">1.0</span><span class="sxs-lookup"><span data-stu-id="3dace-358">1.0</span></span>|<span data-ttu-id="3dace-359">1.7</span><span class="sxs-lookup"><span data-stu-id="3dace-359">1.7</span></span>|
|[<span data-ttu-id="3dace-360">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3dace-360">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="3dace-361">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3dace-361">ReadItem</span></span>|<span data-ttu-id="3dace-362">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="3dace-362">ReadWriteItem</span></span>|
|[<span data-ttu-id="3dace-363">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3dace-363">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="3dace-364">Read</span><span class="sxs-lookup"><span data-stu-id="3dace-364">Read</span></span>|<span data-ttu-id="3dace-365">撰写</span><span class="sxs-lookup"><span data-stu-id="3dace-365">Compose</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="3dace-366">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="3dace-366">internetMessageId :String</span></span>

<span data-ttu-id="3dace-p113">获取电子邮件的 Internet 消息标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="3dace-p113">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="3dace-369">类型：</span><span class="sxs-lookup"><span data-stu-id="3dace-369">Type:</span></span>

*   <span data-ttu-id="3dace-370">String</span><span class="sxs-lookup"><span data-stu-id="3dace-370">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="3dace-371">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-371">Requirements</span></span>

|<span data-ttu-id="3dace-372">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-372">Requirement</span></span>|<span data-ttu-id="3dace-373">值</span><span class="sxs-lookup"><span data-stu-id="3dace-373">Value</span></span>|
|---|---|
|[<span data-ttu-id="3dace-374">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3dace-374">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="3dace-375">1.0</span><span class="sxs-lookup"><span data-stu-id="3dace-375">1.0</span></span>|
|[<span data-ttu-id="3dace-376">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3dace-376">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="3dace-377">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3dace-377">ReadItem</span></span>|
|[<span data-ttu-id="3dace-378">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3dace-378">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="3dace-379">阅读</span><span class="sxs-lookup"><span data-stu-id="3dace-379">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3dace-380">示例</span><span class="sxs-lookup"><span data-stu-id="3dace-380">Example</span></span>

```
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="3dace-381">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="3dace-381">itemClass :String</span></span>

<span data-ttu-id="3dace-p114">获取选定项目的 Exchange Web 服务项目类。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="3dace-p114">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="3dace-p115">`itemClass` 属性指定所选项目的邮件类别。以下是邮件或约会项目的默认邮件类别。</span><span class="sxs-lookup"><span data-stu-id="3dace-p115">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="3dace-386">类型</span><span class="sxs-lookup"><span data-stu-id="3dace-386">Type</span></span>|<span data-ttu-id="3dace-387">说明</span><span class="sxs-lookup"><span data-stu-id="3dace-387">Description</span></span>|<span data-ttu-id="3dace-388">项目类</span><span class="sxs-lookup"><span data-stu-id="3dace-388">item class</span></span>|
|---|---|---|
|<span data-ttu-id="3dace-389">约会项目</span><span class="sxs-lookup"><span data-stu-id="3dace-389">Appointment items</span></span>|<span data-ttu-id="3dace-390">这些是项目类 `IPM.Appointment` 或 `IPM.Appointment.Occurence` 的日历项目。</span><span class="sxs-lookup"><span data-stu-id="3dace-390">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurence`|
|<span data-ttu-id="3dace-391">邮件项目</span><span class="sxs-lookup"><span data-stu-id="3dace-391">Message items</span></span>|<span data-ttu-id="3dace-392">这些项目包括具有默认邮件类别 `IPM.Note` 的电子邮件，以及使用 `IPM.Schedule.Meeting` 作为基础邮件类别的会议请求、响应和取消。</span><span class="sxs-lookup"><span data-stu-id="3dace-392">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="3dace-393">你可以创建用于扩展默认邮件类别的自定义邮件类别，例如，自定义约会邮件类别 `IPM.Appointment.Contoso`。</span><span class="sxs-lookup"><span data-stu-id="3dace-393">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="3dace-394">类型：</span><span class="sxs-lookup"><span data-stu-id="3dace-394">Type:</span></span>

*   <span data-ttu-id="3dace-395">String</span><span class="sxs-lookup"><span data-stu-id="3dace-395">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="3dace-396">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-396">Requirements</span></span>

|<span data-ttu-id="3dace-397">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-397">Requirement</span></span>|<span data-ttu-id="3dace-398">值</span><span class="sxs-lookup"><span data-stu-id="3dace-398">Value</span></span>|
|---|---|
|[<span data-ttu-id="3dace-399">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3dace-399">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="3dace-400">1.0</span><span class="sxs-lookup"><span data-stu-id="3dace-400">1.0</span></span>|
|[<span data-ttu-id="3dace-401">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3dace-401">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="3dace-402">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3dace-402">ReadItem</span></span>|
|[<span data-ttu-id="3dace-403">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3dace-403">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="3dace-404">阅读</span><span class="sxs-lookup"><span data-stu-id="3dace-404">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3dace-405">示例</span><span class="sxs-lookup"><span data-stu-id="3dace-405">Example</span></span>

```
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="3dace-406">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="3dace-406">(nullable) itemId :String</span></span>

<span data-ttu-id="3dace-p116">获取当前项目的 Exchange Web 服务项目标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="3dace-p116">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="3dace-409">`itemId` 属性返回的标识符与 Exchange Web 服务项目标识符相同。</span><span class="sxs-lookup"><span data-stu-id="3dace-409">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="3dace-410">`itemId`属性不是与 Outlook 条目 ID 或使用 Outlook REST API 的 ID。</span><span class="sxs-lookup"><span data-stu-id="3dace-410">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="3dace-411">使用此值的 REST API 调用之前，应将其转换使用[Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string)。</span><span class="sxs-lookup"><span data-stu-id="3dace-411">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="3dace-412">有关详细信息，请参阅[使用 Outlook REST Api 从 Outlook 外接程序](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id)。</span><span class="sxs-lookup"><span data-stu-id="3dace-412">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="3dace-p118">`itemId` 属性在撰写模式下不可用。如果需要项目标识符，[`saveAsync`](#saveasyncoptions-callback) 方法可用于将项目保存到存储，这将在回调函数的 [`AsyncResult.value`](/javascript/api/office/office.asyncresult) 参数中返回项目标识符。</span><span class="sxs-lookup"><span data-stu-id="3dace-p118">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="3dace-415">类型：</span><span class="sxs-lookup"><span data-stu-id="3dace-415">Type:</span></span>

*   <span data-ttu-id="3dace-416">String</span><span class="sxs-lookup"><span data-stu-id="3dace-416">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="3dace-417">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-417">Requirements</span></span>

|<span data-ttu-id="3dace-418">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-418">Requirement</span></span>|<span data-ttu-id="3dace-419">值</span><span class="sxs-lookup"><span data-stu-id="3dace-419">Value</span></span>|
|---|---|
|[<span data-ttu-id="3dace-420">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3dace-420">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="3dace-421">1.0</span><span class="sxs-lookup"><span data-stu-id="3dace-421">1.0</span></span>|
|[<span data-ttu-id="3dace-422">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3dace-422">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="3dace-423">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3dace-423">ReadItem</span></span>|
|[<span data-ttu-id="3dace-424">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3dace-424">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="3dace-425">阅读</span><span class="sxs-lookup"><span data-stu-id="3dace-425">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3dace-426">示例</span><span class="sxs-lookup"><span data-stu-id="3dace-426">Example</span></span>

<span data-ttu-id="3dace-p119">以下代码检查项目标识符是否存在。如果 `itemId` 属性返回 `null` 或 `undefined`，则将项目保存到存储，并从异步结果中获取项目标识符。</span><span class="sxs-lookup"><span data-stu-id="3dace-p119">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook17officemailboxenumsitemtype"></a><span data-ttu-id="3dace-429">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_7/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="3dace-429">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_7/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="3dace-430">获取实例表示的项的类型。</span><span class="sxs-lookup"><span data-stu-id="3dace-430">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="3dace-431">`itemType` 属性返回其中一个 `ItemType` 枚举值，指示 `item` 对象实例是邮件还是约会。</span><span class="sxs-lookup"><span data-stu-id="3dace-431">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="3dace-432">类型：</span><span class="sxs-lookup"><span data-stu-id="3dace-432">Type:</span></span>

*   [<span data-ttu-id="3dace-433">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="3dace-433">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_7/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="3dace-434">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-434">Requirements</span></span>

|<span data-ttu-id="3dace-435">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-435">Requirement</span></span>|<span data-ttu-id="3dace-436">值</span><span class="sxs-lookup"><span data-stu-id="3dace-436">Value</span></span>|
|---|---|
|[<span data-ttu-id="3dace-437">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3dace-437">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="3dace-438">1.0</span><span class="sxs-lookup"><span data-stu-id="3dace-438">1.0</span></span>|
|[<span data-ttu-id="3dace-439">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3dace-439">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="3dace-440">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3dace-440">ReadItem</span></span>|
|[<span data-ttu-id="3dace-441">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3dace-441">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="3dace-442">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3dace-442">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="3dace-443">示例</span><span class="sxs-lookup"><span data-stu-id="3dace-443">Example</span></span>

```
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook17officelocation"></a><span data-ttu-id="3dace-444">location :String|[Location](/javascript/api/outlook_1_7/office.location)</span><span class="sxs-lookup"><span data-stu-id="3dace-444">location :String|[Location](/javascript/api/outlook_1_7/office.location)</span></span>

<span data-ttu-id="3dace-445">获取或设置约会的位置。</span><span class="sxs-lookup"><span data-stu-id="3dace-445">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="3dace-446">阅读模式</span><span class="sxs-lookup"><span data-stu-id="3dace-446">Read mode</span></span>

<span data-ttu-id="3dace-447">`location` 属性返回一个包含约会位置的字符串。</span><span class="sxs-lookup"><span data-stu-id="3dace-447">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="3dace-448">撰写模式</span><span class="sxs-lookup"><span data-stu-id="3dace-448">Compose mode</span></span>

<span data-ttu-id="3dace-449">`location` 属性返回一个 `Location` 对象，该对象提供用于获取和设置约会位置的方法。</span><span class="sxs-lookup"><span data-stu-id="3dace-449">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="3dace-450">类型：</span><span class="sxs-lookup"><span data-stu-id="3dace-450">Type:</span></span>

*   <span data-ttu-id="3dace-451">String | [Location](/javascript/api/outlook_1_7/office.location)</span><span class="sxs-lookup"><span data-stu-id="3dace-451">String | [Location](/javascript/api/outlook_1_7/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="3dace-452">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-452">Requirements</span></span>

|<span data-ttu-id="3dace-453">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-453">Requirement</span></span>|<span data-ttu-id="3dace-454">值</span><span class="sxs-lookup"><span data-stu-id="3dace-454">Value</span></span>|
|---|---|
|[<span data-ttu-id="3dace-455">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3dace-455">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="3dace-456">1.0</span><span class="sxs-lookup"><span data-stu-id="3dace-456">1.0</span></span>|
|[<span data-ttu-id="3dace-457">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3dace-457">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="3dace-458">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3dace-458">ReadItem</span></span>|
|[<span data-ttu-id="3dace-459">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3dace-459">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="3dace-460">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3dace-460">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="3dace-461">示例</span><span class="sxs-lookup"><span data-stu-id="3dace-461">Example</span></span>

```
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="3dace-462">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="3dace-462">normalizedSubject :String</span></span>

<span data-ttu-id="3dace-p120">获取删除了所有前缀（包括 `RE:` 和 `FWD:`）的项目主题。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="3dace-p120">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="3dace-p121">normalizedSubject 属性获取包含由电子邮件程序添加的任何标准前缀（如 `RE:` 和 `FW:`）的项目主题。若要获取包含完整前缀的项目主题，请使用 [`subject`](#subject-stringsubjectjavascriptapioutlook17officesubject) 属性。</span><span class="sxs-lookup"><span data-stu-id="3dace-p121">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook17officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="3dace-467">类型:</span><span class="sxs-lookup"><span data-stu-id="3dace-467">Type:</span></span>

*   <span data-ttu-id="3dace-468">String</span><span class="sxs-lookup"><span data-stu-id="3dace-468">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="3dace-469">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-469">Requirements</span></span>

|<span data-ttu-id="3dace-470">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-470">Requirement</span></span>|<span data-ttu-id="3dace-471">值</span><span class="sxs-lookup"><span data-stu-id="3dace-471">Value</span></span>|
|---|---|
|[<span data-ttu-id="3dace-472">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3dace-472">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="3dace-473">1.0</span><span class="sxs-lookup"><span data-stu-id="3dace-473">1.0</span></span>|
|[<span data-ttu-id="3dace-474">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3dace-474">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="3dace-475">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3dace-475">ReadItem</span></span>|
|[<span data-ttu-id="3dace-476">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3dace-476">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="3dace-477">阅读</span><span class="sxs-lookup"><span data-stu-id="3dace-477">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3dace-478">示例</span><span class="sxs-lookup"><span data-stu-id="3dace-478">Example</span></span>

```
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook17officenotificationmessages"></a><span data-ttu-id="3dace-479">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_7/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="3dace-479">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_7/office.notificationmessages)</span></span>

<span data-ttu-id="3dace-480">获取项目的通知邮件。</span><span class="sxs-lookup"><span data-stu-id="3dace-480">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="3dace-481">类型：</span><span class="sxs-lookup"><span data-stu-id="3dace-481">Type:</span></span>

*   [<span data-ttu-id="3dace-482">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="3dace-482">NotificationMessages</span></span>](/javascript/api/outlook_1_7/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="3dace-483">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-483">Requirements</span></span>

|<span data-ttu-id="3dace-484">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-484">Requirement</span></span>|<span data-ttu-id="3dace-485">值</span><span class="sxs-lookup"><span data-stu-id="3dace-485">Value</span></span>|
|---|---|
|[<span data-ttu-id="3dace-486">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3dace-486">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="3dace-487">1.3</span><span class="sxs-lookup"><span data-stu-id="3dace-487">1.3</span></span>|
|[<span data-ttu-id="3dace-488">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3dace-488">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="3dace-489">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3dace-489">ReadItem</span></span>|
|[<span data-ttu-id="3dace-490">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3dace-490">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="3dace-491">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3dace-491">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="3dace-492">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="3dace-492">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="3dace-493">提供对事件的可选与会者的访问。</span><span class="sxs-lookup"><span data-stu-id="3dace-493">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="3dace-494">对象的类型和访问级别取决于当前项的模式。</span><span class="sxs-lookup"><span data-stu-id="3dace-494">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="3dace-495">阅读模式</span><span class="sxs-lookup"><span data-stu-id="3dace-495">Read mode</span></span>

<span data-ttu-id="3dace-496">`optionalAttendees` 属性返回一个数组，其中包含每个可选与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="3dace-496">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="3dace-497">撰写模式</span><span class="sxs-lookup"><span data-stu-id="3dace-497">Compose mode</span></span>

<span data-ttu-id="3dace-498">`optionalAttendees`属性返回`Recipients`对象，它提供用于获取或更新会议的可选与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="3dace-498">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="3dace-499">类型：</span><span class="sxs-lookup"><span data-stu-id="3dace-499">Type:</span></span>

*   <span data-ttu-id="3dace-500">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="3dace-500">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="3dace-501">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-501">Requirements</span></span>

|<span data-ttu-id="3dace-502">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-502">Requirement</span></span>|<span data-ttu-id="3dace-503">值</span><span class="sxs-lookup"><span data-stu-id="3dace-503">Value</span></span>|
|---|---|
|[<span data-ttu-id="3dace-504">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3dace-504">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="3dace-505">1.0</span><span class="sxs-lookup"><span data-stu-id="3dace-505">1.0</span></span>|
|[<span data-ttu-id="3dace-506">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3dace-506">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="3dace-507">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3dace-507">ReadItem</span></span>|
|[<span data-ttu-id="3dace-508">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3dace-508">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="3dace-509">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3dace-509">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="3dace-510">示例</span><span class="sxs-lookup"><span data-stu-id="3dace-510">Example</span></span>

```
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsorganizerjavascriptapioutlook17officeorganizer"></a><span data-ttu-id="3dace-511">组织者：[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[组织者](/javascript/api/outlook_1_7/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="3dace-511">organizer :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[Organizer](/javascript/api/outlook_1_7/office.organizer)</span></span>

<span data-ttu-id="3dace-512">获取指定会议组织者的电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="3dace-512">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="3dace-513">阅读模式</span><span class="sxs-lookup"><span data-stu-id="3dace-513">Read mode</span></span>

<span data-ttu-id="3dace-514">`organizer`属性返回一个代表会议组织者的[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)对象。</span><span class="sxs-lookup"><span data-stu-id="3dace-514">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="3dace-515">撰写模式</span><span class="sxs-lookup"><span data-stu-id="3dace-515">Compose mode</span></span>

<span data-ttu-id="3dace-516">`organizer`属性返回一个[管理器](/javascript/api/outlook_1_7/office.organizer)对象，提供要获取组织者值的方法。</span><span class="sxs-lookup"><span data-stu-id="3dace-516">The `organizer` property returns an [Organizer](/javascript/api/outlook_1_7/office.organizer) object that provides a method to get the organizer value.</span></span>

##### <a name="type"></a><span data-ttu-id="3dace-517">类型:</span><span class="sxs-lookup"><span data-stu-id="3dace-517">Type:</span></span>

*   <span data-ttu-id="3dace-518">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [组织者](/javascript/api/outlook_1_7/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="3dace-518">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [Organizer](/javascript/api/outlook_1_7/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="3dace-519">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-519">Requirements</span></span>

|<span data-ttu-id="3dace-520">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-520">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="3dace-521">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3dace-521">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="3dace-522">1.0</span><span class="sxs-lookup"><span data-stu-id="3dace-522">1.0</span></span>|<span data-ttu-id="3dace-523">1.7</span><span class="sxs-lookup"><span data-stu-id="3dace-523">1.7</span></span>|
|[<span data-ttu-id="3dace-524">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3dace-524">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="3dace-525">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3dace-525">ReadItem</span></span>|<span data-ttu-id="3dace-526">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="3dace-526">ReadWriteItem</span></span>|
|[<span data-ttu-id="3dace-527">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3dace-527">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="3dace-528">Read</span><span class="sxs-lookup"><span data-stu-id="3dace-528">Read</span></span>|<span data-ttu-id="3dace-529">撰写</span><span class="sxs-lookup"><span data-stu-id="3dace-529">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="3dace-530">示例</span><span class="sxs-lookup"><span data-stu-id="3dace-530">Example</span></span>

```
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

#### <a name="nullable-recurrence-recurrencejavascriptapioutlook17officerecurrence"></a><span data-ttu-id="3dace-531">(可以为 null) 的定期：[定期](/javascript/api/outlook_1_7/office.recurrence)</span><span class="sxs-lookup"><span data-stu-id="3dace-531">(nullable) recurrence :[Recurrence](/javascript/api/outlook_1_7/office.recurrence)</span></span>

<span data-ttu-id="3dace-532">获取或设置约会的定期模式。</span><span class="sxs-lookup"><span data-stu-id="3dace-532">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="3dace-533">获取会议请求的定期模式。</span><span class="sxs-lookup"><span data-stu-id="3dace-533">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="3dace-534">阅读和撰写的约会项目的模式。</span><span class="sxs-lookup"><span data-stu-id="3dace-534">Read and compose modes for appointment items.</span></span> <span data-ttu-id="3dace-535">读取模式的会议请求项目。</span><span class="sxs-lookup"><span data-stu-id="3dace-535">Read mode for meeting request items.</span></span>

<span data-ttu-id="3dace-536">`recurrence`属性返回的定期约会或会议请求的[定期](/javascript/api/outlook_1_7/office.recurrence)对象，如果项是一系列或一系列中的实例。</span><span class="sxs-lookup"><span data-stu-id="3dace-536">The `recurrence` property returns a [recurrence](/javascript/api/outlook_1_7/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="3dace-537">`null`将返回单个约会和会议请求的单个约会。</span><span class="sxs-lookup"><span data-stu-id="3dace-537">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="3dace-538">`undefined`将返回不是会议请求的消息。</span><span class="sxs-lookup"><span data-stu-id="3dace-538">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="3dace-539">注意： 会议请求具有`itemClass`IPM 的值。Schedule.Meeting.Request。</span><span class="sxs-lookup"><span data-stu-id="3dace-539">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="3dace-540">注意： 如果定期对象是`null`，这表明该对象是一个约会或会议请求的一个约会，不属于一系列。</span><span class="sxs-lookup"><span data-stu-id="3dace-540">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="type"></a><span data-ttu-id="3dace-541">类型:</span><span class="sxs-lookup"><span data-stu-id="3dace-541">Type:</span></span>

* [<span data-ttu-id="3dace-542">定期</span><span class="sxs-lookup"><span data-stu-id="3dace-542">Recurrence</span></span>](/javascript/api/outlook_1_7/office.recurrence)

|<span data-ttu-id="3dace-543">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-543">Requirement</span></span>|<span data-ttu-id="3dace-544">值</span><span class="sxs-lookup"><span data-stu-id="3dace-544">Value</span></span>|
|---|---|
|[<span data-ttu-id="3dace-545">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3dace-545">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="3dace-546">1.7</span><span class="sxs-lookup"><span data-stu-id="3dace-546">1.7</span></span>|
|[<span data-ttu-id="3dace-547">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3dace-547">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="3dace-548">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3dace-548">ReadItem</span></span>|
|[<span data-ttu-id="3dace-549">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3dace-549">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="3dace-550">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3dace-550">Compose or read</span></span>|

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="3dace-551">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="3dace-551">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="3dace-552">提供对事件的必需与会者的访问。</span><span class="sxs-lookup"><span data-stu-id="3dace-552">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="3dace-553">对象的类型和访问级别取决于当前项的模式。</span><span class="sxs-lookup"><span data-stu-id="3dace-553">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="3dace-554">阅读模式</span><span class="sxs-lookup"><span data-stu-id="3dace-554">Read mode</span></span>

<span data-ttu-id="3dace-555">`requiredAttendees` 属性返回一个数组，其中包含每个必需与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="3dace-555">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="3dace-556">撰写模式</span><span class="sxs-lookup"><span data-stu-id="3dace-556">Compose mode</span></span>

<span data-ttu-id="3dace-557">`requiredAttendees`属性返回`Recipients`对象，它提供用于获取或更新会议的必需的与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="3dace-557">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="3dace-558">类型：</span><span class="sxs-lookup"><span data-stu-id="3dace-558">Type:</span></span>

*   <span data-ttu-id="3dace-559">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="3dace-559">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="3dace-560">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-560">Requirements</span></span>

|<span data-ttu-id="3dace-561">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-561">Requirement</span></span>|<span data-ttu-id="3dace-562">值</span><span class="sxs-lookup"><span data-stu-id="3dace-562">Value</span></span>|
|---|---|
|[<span data-ttu-id="3dace-563">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3dace-563">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="3dace-564">1.0</span><span class="sxs-lookup"><span data-stu-id="3dace-564">1.0</span></span>|
|[<span data-ttu-id="3dace-565">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3dace-565">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="3dace-566">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3dace-566">ReadItem</span></span>|
|[<span data-ttu-id="3dace-567">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3dace-567">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="3dace-568">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3dace-568">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="3dace-569">示例</span><span class="sxs-lookup"><span data-stu-id="3dace-569">Example</span></span>

```
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetails"></a><span data-ttu-id="3dace-570">sender :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="3dace-570">sender :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)</span></span>

<span data-ttu-id="3dace-p126">获取电子邮件发件人的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="3dace-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="3dace-p127">[`from`](#from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom) 和 `sender` 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="3dace-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="3dace-575">`recipientType`属性`EmailAddressDetails`对象在`sender`属性是`undefined`。</span><span class="sxs-lookup"><span data-stu-id="3dace-575">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="3dace-576">类型:</span><span class="sxs-lookup"><span data-stu-id="3dace-576">Type:</span></span>

*   [<span data-ttu-id="3dace-577">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="3dace-577">EmailAddressDetails</span></span>](/javascript/api/outlook_1_7/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="3dace-578">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-578">Requirements</span></span>

|<span data-ttu-id="3dace-579">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-579">Requirement</span></span>|<span data-ttu-id="3dace-580">值</span><span class="sxs-lookup"><span data-stu-id="3dace-580">Value</span></span>|
|---|---|
|[<span data-ttu-id="3dace-581">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3dace-581">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="3dace-582">1.0</span><span class="sxs-lookup"><span data-stu-id="3dace-582">1.0</span></span>|
|[<span data-ttu-id="3dace-583">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3dace-583">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="3dace-584">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3dace-584">ReadItem</span></span>|
|[<span data-ttu-id="3dace-585">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3dace-585">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="3dace-586">阅读</span><span class="sxs-lookup"><span data-stu-id="3dace-586">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3dace-587">示例</span><span class="sxs-lookup"><span data-stu-id="3dace-587">Example</span></span>

```
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="3dace-588">(nullable) 可能指向： 字符串</span><span class="sxs-lookup"><span data-stu-id="3dace-588">(nullable) seriesId :String</span></span>

<span data-ttu-id="3dace-589">获取实例所属的系列的 id。</span><span class="sxs-lookup"><span data-stu-id="3dace-589">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="3dace-590">在 OWA 和 Outlook 中，`seriesId`返回父 （系列） 项此项目所属的 Exchange Web Services (EWS) ID。</span><span class="sxs-lookup"><span data-stu-id="3dace-590">In OWA and Outlook, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="3dace-591">但是，在 iOS 和 Android，`seriesId`返回父项的 REST ID。</span><span class="sxs-lookup"><span data-stu-id="3dace-591">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="3dace-592">`seriesId` 属性返回的标识符与 Exchange Web 服务项目标识符相同。</span><span class="sxs-lookup"><span data-stu-id="3dace-592">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="3dace-593">`seriesId`属性不是与 Outlook Id 使用 Outlook REST API 相同。</span><span class="sxs-lookup"><span data-stu-id="3dace-593">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="3dace-594">使用此值的 REST API 调用之前，应将其转换使用[Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string)。</span><span class="sxs-lookup"><span data-stu-id="3dace-594">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="3dace-595">有关详细信息，请参阅[使用 Outlook REST Api 从 Outlook 外接程序](https://docs.microsoft.com/outlook/add-ins/use-rest-api)。</span><span class="sxs-lookup"><span data-stu-id="3dace-595">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="3dace-596">`seriesId`属性返回`null`如不具有父项的项单个系列项目、 约会或会议请求，并返回`undefined`不会议请求的其他任何项。</span><span class="sxs-lookup"><span data-stu-id="3dace-596">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="3dace-597">类型:</span><span class="sxs-lookup"><span data-stu-id="3dace-597">Type:</span></span>

* <span data-ttu-id="3dace-598">String</span><span class="sxs-lookup"><span data-stu-id="3dace-598">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="3dace-599">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-599">Requirements</span></span>

|<span data-ttu-id="3dace-600">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-600">Requirement</span></span>|<span data-ttu-id="3dace-601">值</span><span class="sxs-lookup"><span data-stu-id="3dace-601">Value</span></span>|
|---|---|
|[<span data-ttu-id="3dace-602">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3dace-602">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="3dace-603">1.7</span><span class="sxs-lookup"><span data-stu-id="3dace-603">1.7</span></span>|
|[<span data-ttu-id="3dace-604">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3dace-604">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="3dace-605">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3dace-605">ReadItem</span></span>|
|[<span data-ttu-id="3dace-606">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3dace-606">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="3dace-607">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3dace-607">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="3dace-608">示例</span><span class="sxs-lookup"><span data-stu-id="3dace-608">Example</span></span>

```
var seriesId = Office.context.mailbox.item.seriesId; 
var isSeries = (seriesId == null);
```

####  <a name="start-datetimejavascriptapioutlook17officetime"></a><span data-ttu-id="3dace-609">start :Date|[Time](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="3dace-609">start :Date|[Time](/javascript/api/outlook_1_7/office.time)</span></span>

<span data-ttu-id="3dace-610">获取或设置约会开始的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="3dace-610">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="3dace-p130">将 `start` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime) 方法将该值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="3dace-p130">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="3dace-613">阅读模式</span><span class="sxs-lookup"><span data-stu-id="3dace-613">Read mode</span></span>

<span data-ttu-id="3dace-614">`start` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="3dace-614">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="3dace-615">撰写模式</span><span class="sxs-lookup"><span data-stu-id="3dace-615">Compose mode</span></span>

<span data-ttu-id="3dace-616">`start` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="3dace-616">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="3dace-617">使用 [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) 方法设置开始时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="3dace-617">When you use the [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="3dace-618">类型：</span><span class="sxs-lookup"><span data-stu-id="3dace-618">Type:</span></span>

*   <span data-ttu-id="3dace-619">Date | [Time](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="3dace-619">Date | [Time](/javascript/api/outlook_1_7/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="3dace-620">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-620">Requirements</span></span>

|<span data-ttu-id="3dace-621">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-621">Requirement</span></span>|<span data-ttu-id="3dace-622">值</span><span class="sxs-lookup"><span data-stu-id="3dace-622">Value</span></span>|
|---|---|
|[<span data-ttu-id="3dace-623">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3dace-623">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="3dace-624">1.0</span><span class="sxs-lookup"><span data-stu-id="3dace-624">1.0</span></span>|
|[<span data-ttu-id="3dace-625">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3dace-625">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="3dace-626">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3dace-626">ReadItem</span></span>|
|[<span data-ttu-id="3dace-627">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3dace-627">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="3dace-628">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3dace-628">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="3dace-629">示例</span><span class="sxs-lookup"><span data-stu-id="3dace-629">Example</span></span>

<span data-ttu-id="3dace-630">以下示例通过使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) 方法，设置撰写模式下约会的开始时间。</span><span class="sxs-lookup"><span data-stu-id="3dace-630">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```
var startTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used
  // in the callback
     asyncContext: {verb:"Set"}
}
Office.context.mailbox.item.start.setAsync(startTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function
    console.debug("Start Time " + result.asyncContext.verb);
  }
});
```

####  <a name="subject-stringsubjectjavascriptapioutlook17officesubject"></a><span data-ttu-id="3dace-631">subject :String|[Subject](/javascript/api/outlook_1_7/office.subject)</span><span class="sxs-lookup"><span data-stu-id="3dace-631">subject :String|[Subject](/javascript/api/outlook_1_7/office.subject)</span></span>

<span data-ttu-id="3dace-632">获取或设置显示在项目的主题字段中的说明。</span><span class="sxs-lookup"><span data-stu-id="3dace-632">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="3dace-633">`subject` 属性获取或设置由电子邮件服务器发送项目时的整个主题。</span><span class="sxs-lookup"><span data-stu-id="3dace-633">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="3dace-634">阅读模式</span><span class="sxs-lookup"><span data-stu-id="3dace-634">Read mode</span></span>

<span data-ttu-id="3dace-p131">`subject` 属性返回一个字符串。使用 [`normalizedSubject`](#normalizedsubject-string) 属性获取不带任何前导前缀（如 `RE:` 和 `FW:`）的主题。</span><span class="sxs-lookup"><span data-stu-id="3dace-p131">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="3dace-637">撰写模式</span><span class="sxs-lookup"><span data-stu-id="3dace-637">Compose mode</span></span>

<span data-ttu-id="3dace-638">`subject` 属性返回一个 `Subject` 对象，该对象提供用于获取和设置主题的方法。</span><span class="sxs-lookup"><span data-stu-id="3dace-638">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="3dace-639">类型：</span><span class="sxs-lookup"><span data-stu-id="3dace-639">Type:</span></span>

*   <span data-ttu-id="3dace-640">String | [Subject](/javascript/api/outlook_1_7/office.subject)</span><span class="sxs-lookup"><span data-stu-id="3dace-640">String | [Subject](/javascript/api/outlook_1_7/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="3dace-641">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-641">Requirements</span></span>

|<span data-ttu-id="3dace-642">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-642">Requirement</span></span>|<span data-ttu-id="3dace-643">值</span><span class="sxs-lookup"><span data-stu-id="3dace-643">Value</span></span>|
|---|---|
|[<span data-ttu-id="3dace-644">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3dace-644">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="3dace-645">1.0</span><span class="sxs-lookup"><span data-stu-id="3dace-645">1.0</span></span>|
|[<span data-ttu-id="3dace-646">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3dace-646">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="3dace-647">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3dace-647">ReadItem</span></span>|
|[<span data-ttu-id="3dace-648">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3dace-648">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="3dace-649">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3dace-649">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="3dace-650">到： 数组。 <[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[收件人](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="3dace-650">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="3dace-651">提供对邮件的**到**行上的收件人访问。</span><span class="sxs-lookup"><span data-stu-id="3dace-651">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="3dace-652">对象的类型和访问级别取决于当前项的模式。</span><span class="sxs-lookup"><span data-stu-id="3dace-652">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="3dace-653">阅读模式</span><span class="sxs-lookup"><span data-stu-id="3dace-653">Read mode</span></span>

<span data-ttu-id="3dace-p133">`to` 属性返回包含邮件的**收件人**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="3dace-p133">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="3dace-656">撰写模式</span><span class="sxs-lookup"><span data-stu-id="3dace-656">Compose mode</span></span>

<span data-ttu-id="3dace-657">`to`属性返回`Recipients`提供方法用于获取或更新邮件的**到**行上的收件人对象。</span><span class="sxs-lookup"><span data-stu-id="3dace-657">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="3dace-658">类型：</span><span class="sxs-lookup"><span data-stu-id="3dace-658">Type:</span></span>

*   <span data-ttu-id="3dace-659">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="3dace-659">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="3dace-660">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-660">Requirements</span></span>

|<span data-ttu-id="3dace-661">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-661">Requirement</span></span>|<span data-ttu-id="3dace-662">值</span><span class="sxs-lookup"><span data-stu-id="3dace-662">Value</span></span>|
|---|---|
|[<span data-ttu-id="3dace-663">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3dace-663">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="3dace-664">1.0</span><span class="sxs-lookup"><span data-stu-id="3dace-664">1.0</span></span>|
|[<span data-ttu-id="3dace-665">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3dace-665">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="3dace-666">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3dace-666">ReadItem</span></span>|
|[<span data-ttu-id="3dace-667">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3dace-667">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="3dace-668">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3dace-668">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="3dace-669">示例</span><span class="sxs-lookup"><span data-stu-id="3dace-669">Example</span></span>

```
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="3dace-670">方法</span><span class="sxs-lookup"><span data-stu-id="3dace-670">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="3dace-671">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="3dace-671">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="3dace-672">将文件作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="3dace-672">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="3dace-673">`addFileAttachmentAsync` 方法在指定的 URI 上载文件并将其附加到撰写窗体中的项目。</span><span class="sxs-lookup"><span data-stu-id="3dace-673">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="3dace-674">你随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="3dace-674">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3dace-675">参数：</span><span class="sxs-lookup"><span data-stu-id="3dace-675">Parameters:</span></span>
|<span data-ttu-id="3dace-676">名称</span><span class="sxs-lookup"><span data-stu-id="3dace-676">Name</span></span>|<span data-ttu-id="3dace-677">类型</span><span class="sxs-lookup"><span data-stu-id="3dace-677">Type</span></span>|<span data-ttu-id="3dace-678">属性</span><span class="sxs-lookup"><span data-stu-id="3dace-678">Attributes</span></span>|<span data-ttu-id="3dace-679">说明</span><span class="sxs-lookup"><span data-stu-id="3dace-679">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="3dace-680">字符串</span><span class="sxs-lookup"><span data-stu-id="3dace-680">String</span></span>||<span data-ttu-id="3dace-p134">提供附加到邮件或约会的文件的位置的 URI。最大长度为 2048 个字符。</span><span class="sxs-lookup"><span data-stu-id="3dace-p134">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="3dace-683">字符串</span><span class="sxs-lookup"><span data-stu-id="3dace-683">String</span></span>||<span data-ttu-id="3dace-p135">在附件上载过程中显示的附件名称。最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="3dace-p135">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="3dace-686">对象</span><span class="sxs-lookup"><span data-stu-id="3dace-686">Object</span></span>|<span data-ttu-id="3dace-687">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3dace-687">&lt;optional&gt;</span></span>|<span data-ttu-id="3dace-688">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="3dace-688">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="3dace-689">对象</span><span class="sxs-lookup"><span data-stu-id="3dace-689">Object</span></span>|<span data-ttu-id="3dace-690">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3dace-690">&lt;optional&gt;</span></span>|<span data-ttu-id="3dace-691">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="3dace-691">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="3dace-692">布尔值</span><span class="sxs-lookup"><span data-stu-id="3dace-692">Boolean</span></span>|<span data-ttu-id="3dace-693">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3dace-693">&lt;optional&gt;</span></span>|<span data-ttu-id="3dace-694">如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="3dace-694">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="3dace-695">函数</span><span class="sxs-lookup"><span data-stu-id="3dace-695">function</span></span>|<span data-ttu-id="3dace-696">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3dace-696">&lt;optional&gt;</span></span>|<span data-ttu-id="3dace-697">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="3dace-697">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="3dace-698">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="3dace-698">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="3dace-699">如果上传附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="3dace-699">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="3dace-700">错误</span><span class="sxs-lookup"><span data-stu-id="3dace-700">Errors</span></span>

|<span data-ttu-id="3dace-701">错误代码</span><span class="sxs-lookup"><span data-stu-id="3dace-701">Error code</span></span>|<span data-ttu-id="3dace-702">说明</span><span class="sxs-lookup"><span data-stu-id="3dace-702">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="3dace-703">附件大小超过了允许的大小。</span><span class="sxs-lookup"><span data-stu-id="3dace-703">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="3dace-704">该附件的扩展名不是允许的扩展名。</span><span class="sxs-lookup"><span data-stu-id="3dace-704">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="3dace-705">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="3dace-705">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3dace-706">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-706">Requirements</span></span>

|<span data-ttu-id="3dace-707">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-707">Requirement</span></span>|<span data-ttu-id="3dace-708">值</span><span class="sxs-lookup"><span data-stu-id="3dace-708">Value</span></span>|
|---|---|
|[<span data-ttu-id="3dace-709">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3dace-709">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="3dace-710">1.1</span><span class="sxs-lookup"><span data-stu-id="3dace-710">1.1</span></span>|
|[<span data-ttu-id="3dace-711">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3dace-711">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="3dace-712">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="3dace-712">ReadWriteItem</span></span>|
|[<span data-ttu-id="3dace-713">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3dace-713">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="3dace-714">撰写</span><span class="sxs-lookup"><span data-stu-id="3dace-714">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="3dace-715">示例</span><span class="sxs-lookup"><span data-stu-id="3dace-715">Examples</span></span>

```js
function callback(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Attachment added");
  }
}

function addAttachment() {
  // The values in asyncContext can be accessed in the callback
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  var attachmentURL = "https://contoso.com/rtm/icon.png";
  Office.context.mailbox.item.addFileAttachmentAsync(attachmentURL, attachmentURL, options, callback);
}
```

<span data-ttu-id="3dace-716">以下示例将图像文件添加为内联附件，并在邮件正文中引用该附件。</span><span class="sxs-lookup"><span data-stu-id="3dace-716">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

```js
Office.context.mailbox.item.addFileAttachmentAsync
(
  "http://i.imgur.com/WJXklif.png",
  "cute_bird.png",
  {
    isInline: true
  },
  function (asyncResult) {
    Office.context.mailbox.item.body.setAsync(
      "<p>Here's a cute bird!</p><img src='cid:cute_bird.png'>",
      {
        "coercionType": "html"
      },
      function (asyncResult) {
        
      }
    );
  }
);
```

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="3dace-717">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="3dace-717">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="3dace-718">添加支持事件的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="3dace-718">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="3dace-719">当前受支持的事件类型的`Office.EventType.AppointmentTimeChanged`， `Office.EventType.RecipientsChanged`，和`Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="3dace-719">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="3dace-720">参数：</span><span class="sxs-lookup"><span data-stu-id="3dace-720">Parameters:</span></span>

| <span data-ttu-id="3dace-721">名称</span><span class="sxs-lookup"><span data-stu-id="3dace-721">Name</span></span> | <span data-ttu-id="3dace-722">类型</span><span class="sxs-lookup"><span data-stu-id="3dace-722">Type</span></span> | <span data-ttu-id="3dace-723">属性</span><span class="sxs-lookup"><span data-stu-id="3dace-723">Attributes</span></span> | <span data-ttu-id="3dace-724">说明</span><span class="sxs-lookup"><span data-stu-id="3dace-724">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="3dace-725">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="3dace-725">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="3dace-726">应调用处理程序的事件。</span><span class="sxs-lookup"><span data-stu-id="3dace-726">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="3dace-727">函数</span><span class="sxs-lookup"><span data-stu-id="3dace-727">Function</span></span> || <span data-ttu-id="3dace-p136">用于处理事件的函数。此函数必须接受一个参数，即对象文本。参数上的 `type` 属性将匹配传递给 `addHandlerAsync` 的 `eventType` 参数。</span><span class="sxs-lookup"><span data-stu-id="3dace-p136">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="3dace-731">Object</span><span class="sxs-lookup"><span data-stu-id="3dace-731">Object</span></span> | <span data-ttu-id="3dace-732">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3dace-732">&lt;optional&gt;</span></span> | <span data-ttu-id="3dace-733">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="3dace-733">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="3dace-734">对象</span><span class="sxs-lookup"><span data-stu-id="3dace-734">Object</span></span> | <span data-ttu-id="3dace-735">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3dace-735">&lt;optional&gt;</span></span> | <span data-ttu-id="3dace-736">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="3dace-736">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="3dace-737">函数</span><span class="sxs-lookup"><span data-stu-id="3dace-737">function</span></span>| <span data-ttu-id="3dace-738">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3dace-738">&lt;optional&gt;</span></span>|<span data-ttu-id="3dace-739">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="3dace-739">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3dace-740">Requirements</span><span class="sxs-lookup"><span data-stu-id="3dace-740">Requirements</span></span>

|<span data-ttu-id="3dace-741">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-741">Requirement</span></span>| <span data-ttu-id="3dace-742">值</span><span class="sxs-lookup"><span data-stu-id="3dace-742">Value</span></span>|
|---|---|
|[<span data-ttu-id="3dace-743">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3dace-743">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3dace-744">1.7</span><span class="sxs-lookup"><span data-stu-id="3dace-744">1.7</span></span> |
|[<span data-ttu-id="3dace-745">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3dace-745">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3dace-746">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3dace-746">ReadItem</span></span> |
|[<span data-ttu-id="3dace-747">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3dace-747">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3dace-748">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3dace-748">Compose or read</span></span> |

##### <a name="example"></a><span data-ttu-id="3dace-749">示例</span><span class="sxs-lookup"><span data-stu-id="3dace-749">Example</span></span>

```
Office.initialize = function (reason) {
  $(document).ready(function () {
    Office.context.mailbox.item.addHandlerAsync(Office.EventType.RecurrenceChanged, loadNewItem, function (result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        // Handle error
      }
    });
  });
};

function loadNewItem(eventArgs) {
  // Load the properties of the newly selected item
  loadProps(Office.context.mailbox.item);
};
```

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="3dace-750">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="3dace-750">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="3dace-751">将 Exchange 项目（如邮件）作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="3dace-751">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="3dace-p137">`addItemAttachmentAsync` 方法将包含指定 Exchange 标识符的项目附加到撰写窗体中的项目。如果指定一个回调方法，此方法使用 `asyncResult` 参数调用，该参数包含一个附件标识符或代码，指示附加项目过程中出现的任何错误。可以使用 `options` 参数将状态信息传递给回调方法（如果需要）。</span><span class="sxs-lookup"><span data-stu-id="3dace-p137">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="3dace-755">你随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="3dace-755">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="3dace-756">如果您 Office 加载项运行的 Outlook Web App 中`addItemAttachmentAsync`方法可以附加到正在编辑; 的项目之外的项目的项目但是，这不受支持，并且不建议使用。</span><span class="sxs-lookup"><span data-stu-id="3dace-756">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3dace-757">参数：</span><span class="sxs-lookup"><span data-stu-id="3dace-757">Parameters:</span></span>

|<span data-ttu-id="3dace-758">名称</span><span class="sxs-lookup"><span data-stu-id="3dace-758">Name</span></span>|<span data-ttu-id="3dace-759">类型</span><span class="sxs-lookup"><span data-stu-id="3dace-759">Type</span></span>|<span data-ttu-id="3dace-760">属性</span><span class="sxs-lookup"><span data-stu-id="3dace-760">Attributes</span></span>|<span data-ttu-id="3dace-761">说明</span><span class="sxs-lookup"><span data-stu-id="3dace-761">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="3dace-762">字符串</span><span class="sxs-lookup"><span data-stu-id="3dace-762">String</span></span>||<span data-ttu-id="3dace-p138">要附加的项目的 Exchange 标识符。最大长度为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="3dace-p138">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="3dace-765">字符串</span><span class="sxs-lookup"><span data-stu-id="3dace-765">String</span></span>||<span data-ttu-id="3dace-p139">要附加的项目的主题。最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="3dace-p139">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="3dace-768">对象</span><span class="sxs-lookup"><span data-stu-id="3dace-768">Object</span></span>|<span data-ttu-id="3dace-769">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3dace-769">&lt;optional&gt;</span></span>|<span data-ttu-id="3dace-770">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="3dace-770">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="3dace-771">对象</span><span class="sxs-lookup"><span data-stu-id="3dace-771">Object</span></span>|<span data-ttu-id="3dace-772">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3dace-772">&lt;optional&gt;</span></span>|<span data-ttu-id="3dace-773">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="3dace-773">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="3dace-774">函数</span><span class="sxs-lookup"><span data-stu-id="3dace-774">function</span></span>|<span data-ttu-id="3dace-775">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3dace-775">&lt;optional&gt;</span></span>|<span data-ttu-id="3dace-776">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="3dace-776">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="3dace-777">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="3dace-777">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="3dace-778">如果添加附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="3dace-778">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="3dace-779">错误</span><span class="sxs-lookup"><span data-stu-id="3dace-779">Errors</span></span>

|<span data-ttu-id="3dace-780">错误代码</span><span class="sxs-lookup"><span data-stu-id="3dace-780">Error code</span></span>|<span data-ttu-id="3dace-781">说明</span><span class="sxs-lookup"><span data-stu-id="3dace-781">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="3dace-782">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="3dace-782">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3dace-783">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-783">Requirements</span></span>

|<span data-ttu-id="3dace-784">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-784">Requirement</span></span>|<span data-ttu-id="3dace-785">值</span><span class="sxs-lookup"><span data-stu-id="3dace-785">Value</span></span>|
|---|---|
|[<span data-ttu-id="3dace-786">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3dace-786">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="3dace-787">1.1</span><span class="sxs-lookup"><span data-stu-id="3dace-787">1.1</span></span>|
|[<span data-ttu-id="3dace-788">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3dace-788">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="3dace-789">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="3dace-789">ReadWriteItem</span></span>|
|[<span data-ttu-id="3dace-790">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3dace-790">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="3dace-791">撰写</span><span class="sxs-lookup"><span data-stu-id="3dace-791">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="3dace-792">示例</span><span class="sxs-lookup"><span data-stu-id="3dace-792">Example</span></span>

<span data-ttu-id="3dace-793">以下示例将现有的 Outlook 项目添加为名为 `My Attachment` 的附件。</span><span class="sxs-lookup"><span data-stu-id="3dace-793">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

```
function callback(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Attachment added");
  }
}

function addAttachment() {
  // EWS ID of item to attach
  // (Shortened for readability)
  var itemId = "AAMkADI1...AAA=";

  // The values in asyncContext can be accessed in the callback
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  Office.context.mailbox.item.addItemAttachmentAsync(itemId, "My Attachment", options, callback);
}
```

####  <a name="close"></a><span data-ttu-id="3dace-794">close()</span><span class="sxs-lookup"><span data-stu-id="3dace-794">close()</span></span>

<span data-ttu-id="3dace-795">关闭当前正在撰写的项目。</span><span class="sxs-lookup"><span data-stu-id="3dace-795">Closes the current item that is being composed.</span></span>

<span data-ttu-id="3dace-p140">`close` 方法的行为取决于要撰写的项目的当前状态。如果项目具有未保存的更改，客户端将提示用户保存、放弃或取消关闭操作。</span><span class="sxs-lookup"><span data-stu-id="3dace-p140">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="3dace-798">在 Outlook 中的 web，如果该项目是约会和它之前已保存使用在`saveAsync`、 提示用户保存、 放弃，或取消，即使未发生更改由于项目上次保存。</span><span class="sxs-lookup"><span data-stu-id="3dace-798">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="3dace-799">在 Outlook 桌面客户端中，如果邮件是内联答复，`close` 方法不起作用。</span><span class="sxs-lookup"><span data-stu-id="3dace-799">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="3dace-800">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-800">Requirements</span></span>

|<span data-ttu-id="3dace-801">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-801">Requirement</span></span>|<span data-ttu-id="3dace-802">值</span><span class="sxs-lookup"><span data-stu-id="3dace-802">Value</span></span>|
|---|---|
|[<span data-ttu-id="3dace-803">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3dace-803">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="3dace-804">1.3</span><span class="sxs-lookup"><span data-stu-id="3dace-804">1.3</span></span>|
|[<span data-ttu-id="3dace-805">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3dace-805">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="3dace-806">受限</span><span class="sxs-lookup"><span data-stu-id="3dace-806">Restricted</span></span>|
|[<span data-ttu-id="3dace-807">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3dace-807">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="3dace-808">撰写</span><span class="sxs-lookup"><span data-stu-id="3dace-808">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="3dace-809">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="3dace-809">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="3dace-810">显示答复窗体，其中包括所选邮件的发件人和所有收件人或所选约会的组织者和所有与会者。</span><span class="sxs-lookup"><span data-stu-id="3dace-810">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="3dace-811">在适用于 iOS 的 Outlook 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="3dace-811">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="3dace-812">在 Outlook Web App 中，答复窗体显示为包含 3 列视图的弹出式窗体以及包含 2 列或 1 列视图的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="3dace-812">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="3dace-813">如果任意字符串参数超出其限制，`displayReplyAllForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="3dace-813">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="3dace-p141">当在 `formData.attachments` 参数中指定附件时，Outlook 和 Outlook Web App 尝试下载所有附件并将其附加到答复窗体。如果无法添加任何附件，则在窗体 UI 中显示错误。如果这不可能，则不引发错误消息。</span><span class="sxs-lookup"><span data-stu-id="3dace-p141">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3dace-817">参数：</span><span class="sxs-lookup"><span data-stu-id="3dace-817">Parameters:</span></span>

|<span data-ttu-id="3dace-818">名称</span><span class="sxs-lookup"><span data-stu-id="3dace-818">Name</span></span>|<span data-ttu-id="3dace-819">类型</span><span class="sxs-lookup"><span data-stu-id="3dace-819">Type</span></span>|<span data-ttu-id="3dace-820">属性</span><span class="sxs-lookup"><span data-stu-id="3dace-820">Attributes</span></span>|<span data-ttu-id="3dace-821">说明</span><span class="sxs-lookup"><span data-stu-id="3dace-821">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="3dace-822">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="3dace-822">String &#124; Object</span></span>||<span data-ttu-id="3dace-p142">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="3dace-p142">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="3dace-825">**或**</span><span class="sxs-lookup"><span data-stu-id="3dace-825">**OR**</span></span><br/><span data-ttu-id="3dace-p143">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="3dace-p143">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="3dace-828">字符串</span><span class="sxs-lookup"><span data-stu-id="3dace-828">String</span></span>|<span data-ttu-id="3dace-829">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3dace-829">&lt;optional&gt;</span></span>|<span data-ttu-id="3dace-p144">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="3dace-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="3dace-832">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="3dace-832">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="3dace-833">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3dace-833">&lt;optional&gt;</span></span>|<span data-ttu-id="3dace-834">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="3dace-834">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="3dace-835">字符串</span><span class="sxs-lookup"><span data-stu-id="3dace-835">String</span></span>||<span data-ttu-id="3dace-p145">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="3dace-p145">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="3dace-838">字符串</span><span class="sxs-lookup"><span data-stu-id="3dace-838">String</span></span>||<span data-ttu-id="3dace-839">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="3dace-839">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="3dace-840">String</span><span class="sxs-lookup"><span data-stu-id="3dace-840">String</span></span>||<span data-ttu-id="3dace-p146">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="3dace-p146">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="3dace-843">Boolean</span><span class="sxs-lookup"><span data-stu-id="3dace-843">Boolean</span></span>||<span data-ttu-id="3dace-p147">仅在将 `type` 设置为 `file` 时使用。如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="3dace-p147">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="3dace-846">String</span><span class="sxs-lookup"><span data-stu-id="3dace-846">String</span></span>||<span data-ttu-id="3dace-p148">仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="3dace-p148">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="3dace-850">函数</span><span class="sxs-lookup"><span data-stu-id="3dace-850">function</span></span>|<span data-ttu-id="3dace-851">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3dace-851">&lt;optional&gt;</span></span>|<span data-ttu-id="3dace-852">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="3dace-852">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3dace-853">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-853">Requirements</span></span>

|<span data-ttu-id="3dace-854">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-854">Requirement</span></span>|<span data-ttu-id="3dace-855">值</span><span class="sxs-lookup"><span data-stu-id="3dace-855">Value</span></span>|
|---|---|
|[<span data-ttu-id="3dace-856">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3dace-856">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="3dace-857">1.0</span><span class="sxs-lookup"><span data-stu-id="3dace-857">1.0</span></span>|
|[<span data-ttu-id="3dace-858">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3dace-858">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="3dace-859">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3dace-859">ReadItem</span></span>|
|[<span data-ttu-id="3dace-860">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3dace-860">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="3dace-861">阅读</span><span class="sxs-lookup"><span data-stu-id="3dace-861">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="3dace-862">示例</span><span class="sxs-lookup"><span data-stu-id="3dace-862">Examples</span></span>

<span data-ttu-id="3dace-863">以下代码将一个字符串传递到 `displayReplyAllForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="3dace-863">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="3dace-864">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="3dace-864">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="3dace-865">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="3dace-865">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="3dace-866">使用正文和文件附件答复。</span><span class="sxs-lookup"><span data-stu-id="3dace-866">Reply with a body and a file attachment.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    }
  ]
});
```

<span data-ttu-id="3dace-867">使用正文和项目附件答复。</span><span class="sxs-lookup"><span data-stu-id="3dace-867">Reply with a body and an item attachment.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ]
});
```

<span data-ttu-id="3dace-868">使用正文、文件附件、项目附件和回调答复。</span><span class="sxs-lookup"><span data-stu-id="3dace-868">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    },
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ],
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="displayreplyformformdata"></a><span data-ttu-id="3dace-869">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="3dace-869">displayReplyForm(formData)</span></span>

<span data-ttu-id="3dace-870">显示答复窗体，其中仅包括所选邮件的发件人或所选约会的组织者。</span><span class="sxs-lookup"><span data-stu-id="3dace-870">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="3dace-871">在适用于 iOS 的 Outlook 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="3dace-871">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="3dace-872">在 Outlook Web App 中，答复窗体显示为包含 3 列视图的弹出式窗体以及包含 2 列或 1 列视图的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="3dace-872">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="3dace-873">如果任意字符串参数超出其限制，`displayReplyForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="3dace-873">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="3dace-p149">当在 `formData.attachments` 参数中指定附件时，Outlook 和 Outlook Web App 尝试下载所有附件并将其附加到答复窗体。如果无法添加任何附件，则在窗体 UI 中显示错误。如果这不可能，则不引发错误消息。</span><span class="sxs-lookup"><span data-stu-id="3dace-p149">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3dace-877">参数：</span><span class="sxs-lookup"><span data-stu-id="3dace-877">Parameters:</span></span>

|<span data-ttu-id="3dace-878">名称</span><span class="sxs-lookup"><span data-stu-id="3dace-878">Name</span></span>|<span data-ttu-id="3dace-879">类型</span><span class="sxs-lookup"><span data-stu-id="3dace-879">Type</span></span>|<span data-ttu-id="3dace-880">属性</span><span class="sxs-lookup"><span data-stu-id="3dace-880">Attributes</span></span>|<span data-ttu-id="3dace-881">说明</span><span class="sxs-lookup"><span data-stu-id="3dace-881">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="3dace-882">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="3dace-882">String &#124; Object</span></span>||<span data-ttu-id="3dace-p150">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="3dace-p150">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="3dace-885">**或**</span><span class="sxs-lookup"><span data-stu-id="3dace-885">**OR**</span></span><br/><span data-ttu-id="3dace-p151">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="3dace-p151">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="3dace-888">字符串</span><span class="sxs-lookup"><span data-stu-id="3dace-888">String</span></span>|<span data-ttu-id="3dace-889">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3dace-889">&lt;optional&gt;</span></span>|<span data-ttu-id="3dace-p152">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="3dace-p152">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="3dace-892">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="3dace-892">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="3dace-893">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3dace-893">&lt;optional&gt;</span></span>|<span data-ttu-id="3dace-894">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="3dace-894">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="3dace-895">字符串</span><span class="sxs-lookup"><span data-stu-id="3dace-895">String</span></span>||<span data-ttu-id="3dace-p153">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="3dace-p153">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="3dace-898">字符串</span><span class="sxs-lookup"><span data-stu-id="3dace-898">String</span></span>||<span data-ttu-id="3dace-899">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="3dace-899">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="3dace-900">String</span><span class="sxs-lookup"><span data-stu-id="3dace-900">String</span></span>||<span data-ttu-id="3dace-p154">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="3dace-p154">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="3dace-903">Boolean</span><span class="sxs-lookup"><span data-stu-id="3dace-903">Boolean</span></span>||<span data-ttu-id="3dace-p155">仅在将 `type` 设置为 `file` 时使用。如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="3dace-p155">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="3dace-906">String</span><span class="sxs-lookup"><span data-stu-id="3dace-906">String</span></span>||<span data-ttu-id="3dace-p156">仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="3dace-p156">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="3dace-910">函数</span><span class="sxs-lookup"><span data-stu-id="3dace-910">function</span></span>|<span data-ttu-id="3dace-911">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3dace-911">&lt;optional&gt;</span></span>|<span data-ttu-id="3dace-912">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="3dace-912">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3dace-913">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-913">Requirements</span></span>

|<span data-ttu-id="3dace-914">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-914">Requirement</span></span>|<span data-ttu-id="3dace-915">值</span><span class="sxs-lookup"><span data-stu-id="3dace-915">Value</span></span>|
|---|---|
|[<span data-ttu-id="3dace-916">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3dace-916">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="3dace-917">1.0</span><span class="sxs-lookup"><span data-stu-id="3dace-917">1.0</span></span>|
|[<span data-ttu-id="3dace-918">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3dace-918">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="3dace-919">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3dace-919">ReadItem</span></span>|
|[<span data-ttu-id="3dace-920">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3dace-920">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="3dace-921">阅读</span><span class="sxs-lookup"><span data-stu-id="3dace-921">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="3dace-922">示例</span><span class="sxs-lookup"><span data-stu-id="3dace-922">Examples</span></span>

<span data-ttu-id="3dace-923">以下代码将一个字符串传递到 `displayReplyForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="3dace-923">The following code passes a string to the `displayReplyForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="3dace-924">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="3dace-924">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="3dace-925">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="3dace-925">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="3dace-926">使用正文和文件附件答复。</span><span class="sxs-lookup"><span data-stu-id="3dace-926">Reply with a body and a file attachment.</span></span>

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    }
  ]
});
```

<span data-ttu-id="3dace-927">使用正文和项目附件答复。</span><span class="sxs-lookup"><span data-stu-id="3dace-927">Reply with a body and an item attachment.</span></span>

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ]
});
```

<span data-ttu-id="3dace-928">使用正文、文件附件、项目附件和回调答复。</span><span class="sxs-lookup"><span data-stu-id="3dace-928">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    },
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ],
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="getentities--entitiesjavascriptapioutlook17officeentities"></a><span data-ttu-id="3dace-929">getEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="3dace-929">getEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span></span>

<span data-ttu-id="3dace-930">获取在选定的项的正文中找到的实体。</span><span class="sxs-lookup"><span data-stu-id="3dace-930">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="3dace-931">在适用于 iOS 的 Outlook 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="3dace-931">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="3dace-932">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-932">Requirements</span></span>

|<span data-ttu-id="3dace-933">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-933">Requirement</span></span>|<span data-ttu-id="3dace-934">值</span><span class="sxs-lookup"><span data-stu-id="3dace-934">Value</span></span>|
|---|---|
|[<span data-ttu-id="3dace-935">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3dace-935">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="3dace-936">1.0</span><span class="sxs-lookup"><span data-stu-id="3dace-936">1.0</span></span>|
|[<span data-ttu-id="3dace-937">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3dace-937">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="3dace-938">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3dace-938">ReadItem</span></span>|
|[<span data-ttu-id="3dace-939">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3dace-939">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="3dace-940">阅读</span><span class="sxs-lookup"><span data-stu-id="3dace-940">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="3dace-941">返回：</span><span class="sxs-lookup"><span data-stu-id="3dace-941">Returns:</span></span>

<span data-ttu-id="3dace-942">类型：[Entities](/javascript/api/outlook_1_7/office.entities)</span><span class="sxs-lookup"><span data-stu-id="3dace-942">Type: [Entities](/javascript/api/outlook_1_7/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="3dace-943">示例</span><span class="sxs-lookup"><span data-stu-id="3dace-943">Example</span></span>

<span data-ttu-id="3dace-944">以下示例访问当前项的主体中的联系人实体。</span><span class="sxs-lookup"><span data-stu-id="3dace-944">The following example accesses the contacts entities in the current item's body.</span></span>

```
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion"></a><span data-ttu-id="3dace-945">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="3dace-945">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span></span>

<span data-ttu-id="3dace-946">获取选定的项的正文中找到的指定的实体类型的所有实体的数组。</span><span class="sxs-lookup"><span data-stu-id="3dace-946">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="3dace-947">在适用于 iOS 的 Outlook 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="3dace-947">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3dace-948">参数：</span><span class="sxs-lookup"><span data-stu-id="3dace-948">Parameters:</span></span>

|<span data-ttu-id="3dace-949">名称</span><span class="sxs-lookup"><span data-stu-id="3dace-949">Name</span></span>|<span data-ttu-id="3dace-950">类型</span><span class="sxs-lookup"><span data-stu-id="3dace-950">Type</span></span>|<span data-ttu-id="3dace-951">说明</span><span class="sxs-lookup"><span data-stu-id="3dace-951">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="3dace-952">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="3dace-952">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_7/office.mailboxenums.entitytype)|<span data-ttu-id="3dace-953">EntityType 枚举值之一。</span><span class="sxs-lookup"><span data-stu-id="3dace-953">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3dace-954">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-954">Requirements</span></span>

|<span data-ttu-id="3dace-955">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-955">Requirement</span></span>|<span data-ttu-id="3dace-956">值</span><span class="sxs-lookup"><span data-stu-id="3dace-956">Value</span></span>|
|---|---|
|[<span data-ttu-id="3dace-957">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3dace-957">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="3dace-958">1.0</span><span class="sxs-lookup"><span data-stu-id="3dace-958">1.0</span></span>|
|[<span data-ttu-id="3dace-959">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3dace-959">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="3dace-960">受限</span><span class="sxs-lookup"><span data-stu-id="3dace-960">Restricted</span></span>|
|[<span data-ttu-id="3dace-961">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3dace-961">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="3dace-962">阅读</span><span class="sxs-lookup"><span data-stu-id="3dace-962">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="3dace-963">返回：</span><span class="sxs-lookup"><span data-stu-id="3dace-963">Returns:</span></span>

<span data-ttu-id="3dace-964">如果在 `entityType` 中传递的值不是 `EntityType` 枚举的有效成员，该方法返回 null。</span><span class="sxs-lookup"><span data-stu-id="3dace-964">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="3dace-965">如果没有指定类型的实体的存在于项目的正文中，该方法将返回空数组。</span><span class="sxs-lookup"><span data-stu-id="3dace-965">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="3dace-966">否则，返回的数组中对象的类型取决于 `entityType` 参数中请求实体的类型。</span><span class="sxs-lookup"><span data-stu-id="3dace-966">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="3dace-967">当使用此方法的最低权限级别**受限**时，某些实体类型需要 **ReadItem** 才能进行访问，如下表中所指定。</span><span class="sxs-lookup"><span data-stu-id="3dace-967">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="3dace-968">`entityType` 的值</span><span class="sxs-lookup"><span data-stu-id="3dace-968">Value of `entityType`</span></span>|<span data-ttu-id="3dace-969">返回的数组中对象的类型</span><span class="sxs-lookup"><span data-stu-id="3dace-969">Type of objects in returned array</span></span>|<span data-ttu-id="3dace-970">所需权限级别</span><span class="sxs-lookup"><span data-stu-id="3dace-970">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="3dace-971">String</span><span class="sxs-lookup"><span data-stu-id="3dace-971">String</span></span>|<span data-ttu-id="3dace-972">**受限**</span><span class="sxs-lookup"><span data-stu-id="3dace-972">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="3dace-973">Contact</span><span class="sxs-lookup"><span data-stu-id="3dace-973">Contact</span></span>|<span data-ttu-id="3dace-974">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="3dace-974">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="3dace-975">String</span><span class="sxs-lookup"><span data-stu-id="3dace-975">String</span></span>|<span data-ttu-id="3dace-976">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="3dace-976">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="3dace-977">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="3dace-977">MeetingSuggestion</span></span>|<span data-ttu-id="3dace-978">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="3dace-978">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="3dace-979">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="3dace-979">PhoneNumber</span></span>|<span data-ttu-id="3dace-980">**受限**</span><span class="sxs-lookup"><span data-stu-id="3dace-980">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="3dace-981">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="3dace-981">TaskSuggestion</span></span>|<span data-ttu-id="3dace-982">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="3dace-982">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="3dace-983">String</span><span class="sxs-lookup"><span data-stu-id="3dace-983">String</span></span>|<span data-ttu-id="3dace-984">**受限**</span><span class="sxs-lookup"><span data-stu-id="3dace-984">**Restricted**</span></span>|

<span data-ttu-id="3dace-985">类型：Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="3dace-985">Type: Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="3dace-986">示例</span><span class="sxs-lookup"><span data-stu-id="3dace-986">Example</span></span>

<span data-ttu-id="3dace-987">下面的示例演示如何访问表示当前项目的正文中的邮政地址的字符串数组。</span><span class="sxs-lookup"><span data-stu-id="3dace-987">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

```
// The initialize function is required for all apps.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    var item = Office.context.mailbox.item;
    // Get an array of strings that represent postal addresses in the current item's body.
    var addresses = item.getEntitiesByType(Office.MailboxEnums.EntityType.Address);
    // Continue processing the array of addresses.
  });
}
```

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion"></a><span data-ttu-id="3dace-988">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="3dace-988">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span></span>

<span data-ttu-id="3dace-989">返回传递清单 XML 文件中定义的命名筛选器的所选项目中的已知实体。</span><span class="sxs-lookup"><span data-stu-id="3dace-989">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="3dace-990">在适用于 iOS 的 Outlook 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="3dace-990">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="3dace-991">`getFilteredEntitiesByName` 方法返回匹配在具有指定 `FilterName` 元素值的清单 XML 文件中的 [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) 规则元素中定义的正则表达式的实体。</span><span class="sxs-lookup"><span data-stu-id="3dace-991">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3dace-992">参数：</span><span class="sxs-lookup"><span data-stu-id="3dace-992">Parameters:</span></span>

|<span data-ttu-id="3dace-993">名称</span><span class="sxs-lookup"><span data-stu-id="3dace-993">Name</span></span>|<span data-ttu-id="3dace-994">类型</span><span class="sxs-lookup"><span data-stu-id="3dace-994">Type</span></span>|<span data-ttu-id="3dace-995">说明</span><span class="sxs-lookup"><span data-stu-id="3dace-995">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="3dace-996">字符串</span><span class="sxs-lookup"><span data-stu-id="3dace-996">String</span></span>|<span data-ttu-id="3dace-997">定义筛选器匹配的 `ItemHasKnownEntity` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="3dace-997">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3dace-998">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-998">Requirements</span></span>

|<span data-ttu-id="3dace-999">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-999">Requirement</span></span>|<span data-ttu-id="3dace-1000">值</span><span class="sxs-lookup"><span data-stu-id="3dace-1000">Value</span></span>|
|---|---|
|[<span data-ttu-id="3dace-1001">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3dace-1001">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="3dace-1002">1.0</span><span class="sxs-lookup"><span data-stu-id="3dace-1002">1.0</span></span>|
|[<span data-ttu-id="3dace-1003">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3dace-1003">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="3dace-1004">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3dace-1004">ReadItem</span></span>|
|[<span data-ttu-id="3dace-1005">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3dace-1005">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="3dace-1006">阅读</span><span class="sxs-lookup"><span data-stu-id="3dace-1006">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="3dace-1007">返回：</span><span class="sxs-lookup"><span data-stu-id="3dace-1007">Returns:</span></span>

<span data-ttu-id="3dace-p158">如果具有匹配 `name` 参数的 `FilterName` 元素值的清单中没有任何 `ItemHasKnownEntity` 元素，则该方法返回 `null`。如果 `name` 参数匹配清单中的 `ItemHasKnownEntity` 元素，但在匹配的当前项目中没有实体，则该方法返回一个空数组。</span><span class="sxs-lookup"><span data-stu-id="3dace-p158">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="3dace-1010">类型：Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="3dace-1010">Type: Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="3dace-1011">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="3dace-1011">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="3dace-1012">返回所选项目中匹配在清单 XML 文件中定义的正则表达式的字符串值。</span><span class="sxs-lookup"><span data-stu-id="3dace-1012">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="3dace-1013">在适用于 iOS 的 Outlook 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="3dace-1013">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="3dace-p159">`getRegExMatches` 方法返回匹配在清单 XML 文件中的每个 `ItemHasRegularExpressionMatch` 或 `ItemHasKnownEntity` 规则元素中定义的正则表达式的字符串。对于 `ItemHasRegularExpressionMatch` 规则，匹配字符串必须发生在该规则指定的项目的属性中。`PropertyName` 简单类型定义支持的属性。</span><span class="sxs-lookup"><span data-stu-id="3dace-p159">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="3dace-1017">例如，考虑一个外接程序清单具有以下 `Rule` 元素：</span><span class="sxs-lookup"><span data-stu-id="3dace-1017">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="3dace-1018">从 `getRegExMatches` 返回的对象应有两个属性：`fruits` 和 `veggies`。</span><span class="sxs-lookup"><span data-stu-id="3dace-1018">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="3dace-p160">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。而是使用 [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) 方法检索整个正文。</span><span class="sxs-lookup"><span data-stu-id="3dace-p160">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="3dace-1022">Requirements</span><span class="sxs-lookup"><span data-stu-id="3dace-1022">Requirements</span></span>

|<span data-ttu-id="3dace-1023">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-1023">Requirement</span></span>|<span data-ttu-id="3dace-1024">值</span><span class="sxs-lookup"><span data-stu-id="3dace-1024">Value</span></span>|
|---|---|
|[<span data-ttu-id="3dace-1025">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3dace-1025">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="3dace-1026">1.0</span><span class="sxs-lookup"><span data-stu-id="3dace-1026">1.0</span></span>|
|[<span data-ttu-id="3dace-1027">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3dace-1027">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="3dace-1028">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3dace-1028">ReadItem</span></span>|
|[<span data-ttu-id="3dace-1029">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3dace-1029">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="3dace-1030">阅读</span><span class="sxs-lookup"><span data-stu-id="3dace-1030">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="3dace-1031">返回：</span><span class="sxs-lookup"><span data-stu-id="3dace-1031">Returns:</span></span>

<span data-ttu-id="3dace-p161">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串数组的对象。每个数组的名称等于匹配 `ItemHasRegularExpressionMatch` 规则的 `RegExName` 属性或匹配 `ItemHasKnownEntity` 规则的 `FilterName` 属性的相应值。</span><span class="sxs-lookup"><span data-stu-id="3dace-p161">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="3dace-1034">

<dt>
类型</dt>


</span><span class="sxs-lookup"><span data-stu-id="3dace-1034">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="3dace-1035">对象</span><span class="sxs-lookup"><span data-stu-id="3dace-1035">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="3dace-1036">示例</span><span class="sxs-lookup"><span data-stu-id="3dace-1036">Example</span></span>

<span data-ttu-id="3dace-1037">以下示例显示了如何访问正则表达式规则元素 `fruits` 和 `veggies` 的匹配项的数组，这些元素在清单中指定。</span><span class="sxs-lookup"><span data-stu-id="3dace-1037">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="3dace-1038">getregexmatchesbyname （name） → (nullable) {数组。 < 字符串 >}</span><span class="sxs-lookup"><span data-stu-id="3dace-1038">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="3dace-1039">返回匹配在清单 XML 文件中定义的命名正则表达式的所选项目中的字符串值。</span><span class="sxs-lookup"><span data-stu-id="3dace-1039">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="3dace-1040">在适用于 iOS 的 Outlook 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="3dace-1040">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="3dace-1041">`getRegExMatchesByName` 方法返回匹配在具有指定 `RegExName` 元素值的清单 XML 文件中的 `ItemHasRegularExpressionMatch` 规则元素中定义的正则表达式的字符串。</span><span class="sxs-lookup"><span data-stu-id="3dace-1041">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="3dace-p162">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。</span><span class="sxs-lookup"><span data-stu-id="3dace-p162">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3dace-1044">参数：</span><span class="sxs-lookup"><span data-stu-id="3dace-1044">Parameters:</span></span>

|<span data-ttu-id="3dace-1045">名称</span><span class="sxs-lookup"><span data-stu-id="3dace-1045">Name</span></span>|<span data-ttu-id="3dace-1046">类型</span><span class="sxs-lookup"><span data-stu-id="3dace-1046">Type</span></span>|<span data-ttu-id="3dace-1047">说明</span><span class="sxs-lookup"><span data-stu-id="3dace-1047">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="3dace-1048">字符串</span><span class="sxs-lookup"><span data-stu-id="3dace-1048">String</span></span>|<span data-ttu-id="3dace-1049">定义筛选器匹配的 `ItemHasRegularExpressionMatch` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="3dace-1049">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3dace-1050">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-1050">Requirements</span></span>

|<span data-ttu-id="3dace-1051">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-1051">Requirement</span></span>|<span data-ttu-id="3dace-1052">值</span><span class="sxs-lookup"><span data-stu-id="3dace-1052">Value</span></span>|
|---|---|
|[<span data-ttu-id="3dace-1053">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3dace-1053">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="3dace-1054">1.0</span><span class="sxs-lookup"><span data-stu-id="3dace-1054">1.0</span></span>|
|[<span data-ttu-id="3dace-1055">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3dace-1055">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="3dace-1056">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3dace-1056">ReadItem</span></span>|
|[<span data-ttu-id="3dace-1057">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3dace-1057">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="3dace-1058">阅读</span><span class="sxs-lookup"><span data-stu-id="3dace-1058">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="3dace-1059">返回：</span><span class="sxs-lookup"><span data-stu-id="3dace-1059">Returns:</span></span>

<span data-ttu-id="3dace-1060">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串的数组。</span><span class="sxs-lookup"><span data-stu-id="3dace-1060">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="3dace-1061">

<dt>类型</dt>

</span><span class="sxs-lookup"><span data-stu-id="3dace-1061">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="3dace-1062">数组。 < 字符串 ></span><span class="sxs-lookup"><span data-stu-id="3dace-1062">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="3dace-1063">示例</span><span class="sxs-lookup"><span data-stu-id="3dace-1063">Example</span></span>

```
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="3dace-1064">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="3dace-1064">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="3dace-1065">以异步方式返回邮件的主题或正文中选定的数据。</span><span class="sxs-lookup"><span data-stu-id="3dace-1065">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="3dace-p163">如果没有选定内容，但光标位于正文或主题中，此方法将会为所选数据返回 null。如果选定的是字段，而不是正文或主题，则此方法返回 `InvalidSelection` 错误。</span><span class="sxs-lookup"><span data-stu-id="3dace-p163">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3dace-1068">参数：</span><span class="sxs-lookup"><span data-stu-id="3dace-1068">Parameters:</span></span>

|<span data-ttu-id="3dace-1069">名称</span><span class="sxs-lookup"><span data-stu-id="3dace-1069">Name</span></span>|<span data-ttu-id="3dace-1070">类型</span><span class="sxs-lookup"><span data-stu-id="3dace-1070">Type</span></span>|<span data-ttu-id="3dace-1071">属性</span><span class="sxs-lookup"><span data-stu-id="3dace-1071">Attributes</span></span>|<span data-ttu-id="3dace-1072">说明</span><span class="sxs-lookup"><span data-stu-id="3dace-1072">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="3dace-1073">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="3dace-1073">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="3dace-p164">请求数据的格式。如果为文本，则此方法返回纯文本作为字符串，删除任何显示的 HTML 标记。如果为 HTML，则此方法返回所选文本，不论是纯文本还是 HTML。</span><span class="sxs-lookup"><span data-stu-id="3dace-p164">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="3dace-1077">对象</span><span class="sxs-lookup"><span data-stu-id="3dace-1077">Object</span></span>|<span data-ttu-id="3dace-1078">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3dace-1078">&lt;optional&gt;</span></span>|<span data-ttu-id="3dace-1079">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="3dace-1079">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="3dace-1080">对象</span><span class="sxs-lookup"><span data-stu-id="3dace-1080">Object</span></span>|<span data-ttu-id="3dace-1081">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3dace-1081">&lt;optional&gt;</span></span>|<span data-ttu-id="3dace-1082">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="3dace-1082">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="3dace-1083">函数</span><span class="sxs-lookup"><span data-stu-id="3dace-1083">function</span></span>||<span data-ttu-id="3dace-1084">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="3dace-1084">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="3dace-1085">若要从回调方法访问所选数据，请调用 `asyncResult.value.data`。</span><span class="sxs-lookup"><span data-stu-id="3dace-1085">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="3dace-1086">若要访问所选内容来自源属性，请调用`asyncResult.value.sourceProperty`，其中将为`body`或`subject`。</span><span class="sxs-lookup"><span data-stu-id="3dace-1086">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3dace-1087">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-1087">Requirements</span></span>

|<span data-ttu-id="3dace-1088">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-1088">Requirement</span></span>|<span data-ttu-id="3dace-1089">值</span><span class="sxs-lookup"><span data-stu-id="3dace-1089">Value</span></span>|
|---|---|
|[<span data-ttu-id="3dace-1090">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3dace-1090">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="3dace-1091">1.2</span><span class="sxs-lookup"><span data-stu-id="3dace-1091">1.2</span></span>|
|[<span data-ttu-id="3dace-1092">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3dace-1092">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="3dace-1093">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="3dace-1093">ReadWriteItem</span></span>|
|[<span data-ttu-id="3dace-1094">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3dace-1094">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="3dace-1095">撰写</span><span class="sxs-lookup"><span data-stu-id="3dace-1095">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="3dace-1096">返回：</span><span class="sxs-lookup"><span data-stu-id="3dace-1096">Returns:</span></span>

<span data-ttu-id="3dace-1097">作为字符串的所选数据的格式由 `coercionType` 确定。</span><span class="sxs-lookup"><span data-stu-id="3dace-1097">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="3dace-1098">

<dt>
类型</dt>


</span><span class="sxs-lookup"><span data-stu-id="3dace-1098">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="3dace-1099">字符串</span><span class="sxs-lookup"><span data-stu-id="3dace-1099">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="3dace-1100">示例</span><span class="sxs-lookup"><span data-stu-id="3dace-1100">Example</span></span>

```
// getting selected data
Office.initialize = function () {
    Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Text, {}, getCallback);
}

function getCallback(asyncResult) {
    var text = asyncResult.value.data;
    var prop = asyncResult.value.sourceProperty;

    Office.context.mailbox.item.setSelectedDataAsync('Setting ' + prop + ': ' + text, {}, setCallback);
}

function setCallback(asyncResult) {
    // check for errors
}
```

#### <a name="getselectedentities--entitiesjavascriptapioutlook17officeentities"></a><span data-ttu-id="3dace-1101">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="3dace-1101">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span></span>

<span data-ttu-id="3dace-p166">获取在用户已选择的突出显示匹配项中找到的实体。突出显示匹配项适用于[上下文外接程序](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins)。</span><span class="sxs-lookup"><span data-stu-id="3dace-p166">Gets the entities found in a highlighted match a user has selected. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="3dace-1104">在适用于 iOS 的 Outlook 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="3dace-1104">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="3dace-1105">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-1105">Requirements</span></span>

|<span data-ttu-id="3dace-1106">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-1106">Requirement</span></span>|<span data-ttu-id="3dace-1107">值</span><span class="sxs-lookup"><span data-stu-id="3dace-1107">Value</span></span>|
|---|---|
|[<span data-ttu-id="3dace-1108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3dace-1108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="3dace-1109">1.6</span><span class="sxs-lookup"><span data-stu-id="3dace-1109">1.6</span></span>|
|[<span data-ttu-id="3dace-1110">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3dace-1110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="3dace-1111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3dace-1111">ReadItem</span></span>|
|[<span data-ttu-id="3dace-1112">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3dace-1112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="3dace-1113">阅读</span><span class="sxs-lookup"><span data-stu-id="3dace-1113">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="3dace-1114">返回：</span><span class="sxs-lookup"><span data-stu-id="3dace-1114">Returns:</span></span>

<span data-ttu-id="3dace-1115">类型：[Entities](/javascript/api/outlook_1_7/office.entities)</span><span class="sxs-lookup"><span data-stu-id="3dace-1115">Type: [Entities](/javascript/api/outlook_1_7/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="3dace-1116">示例</span><span class="sxs-lookup"><span data-stu-id="3dace-1116">Example</span></span>

<span data-ttu-id="3dace-1117">以下示例访问用户选择的突出显示匹配项中的地址实体。</span><span class="sxs-lookup"><span data-stu-id="3dace-1117">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="3dace-1118">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="3dace-1118">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="3dace-p167">返回突出显示匹配项中匹配在清单 XML 文件中定义的正则表达式的字符串值。突出显示匹配项适用于[上下文外接程序](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins)。</span><span class="sxs-lookup"><span data-stu-id="3dace-p167">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="3dace-1121">在适用于 iOS 的 Outlook 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="3dace-1121">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="3dace-p168">`getSelectedRegExMatches` 方法返回匹配在清单 XML 文件中的每个 `ItemHasRegularExpressionMatch` 或 `ItemHasKnownEntity` 规则元素中定义的正则表达式的字符串。对于 `ItemHasRegularExpressionMatch` 规则，匹配字符串必须发生在该规则指定的项目的属性中。`PropertyName` 简单类型定义支持的属性。</span><span class="sxs-lookup"><span data-stu-id="3dace-p168">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="3dace-1125">例如，考虑一个外接程序清单具有以下 `Rule` 元素：</span><span class="sxs-lookup"><span data-stu-id="3dace-1125">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="3dace-1126">从 `getRegExMatches` 返回的对象应有两个属性：`fruits` 和 `veggies`。</span><span class="sxs-lookup"><span data-stu-id="3dace-1126">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="3dace-p169">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。而是使用 [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) 方法检索整个正文。</span><span class="sxs-lookup"><span data-stu-id="3dace-p169">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="3dace-1130">Requirements</span><span class="sxs-lookup"><span data-stu-id="3dace-1130">Requirements</span></span>

|<span data-ttu-id="3dace-1131">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-1131">Requirement</span></span>|<span data-ttu-id="3dace-1132">值</span><span class="sxs-lookup"><span data-stu-id="3dace-1132">Value</span></span>|
|---|---|
|[<span data-ttu-id="3dace-1133">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3dace-1133">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="3dace-1134">1.6</span><span class="sxs-lookup"><span data-stu-id="3dace-1134">1.6</span></span>|
|[<span data-ttu-id="3dace-1135">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3dace-1135">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="3dace-1136">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3dace-1136">ReadItem</span></span>|
|[<span data-ttu-id="3dace-1137">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3dace-1137">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="3dace-1138">阅读</span><span class="sxs-lookup"><span data-stu-id="3dace-1138">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="3dace-1139">返回：</span><span class="sxs-lookup"><span data-stu-id="3dace-1139">Returns:</span></span>

<span data-ttu-id="3dace-p170">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串数组的对象。每个数组的名称等于匹配 `ItemHasRegularExpressionMatch` 规则的 `RegExName` 属性或匹配 `ItemHasKnownEntity` 规则的 `FilterName` 属性的相应值。</span><span class="sxs-lookup"><span data-stu-id="3dace-p170">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="3dace-1142">示例</span><span class="sxs-lookup"><span data-stu-id="3dace-1142">Example</span></span>

<span data-ttu-id="3dace-1143">以下示例显示了如何访问正则表达式规则元素 `fruits` 和 `veggies` 的匹配项的数组，这些元素在清单中指定。</span><span class="sxs-lookup"><span data-stu-id="3dace-1143">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="3dace-1144">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="3dace-1144">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="3dace-1145">异步加载所选项目上此外接程序的自定义属性。</span><span class="sxs-lookup"><span data-stu-id="3dace-1145">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="3dace-p171">自定义属性基于每个应用、每个项目存储为键/值对。此方法在回调中返回 `CustomProperties` 对象，该回调提供访问特定于当前项目和当前外接程序的自定义属性的方法。自定义属性未在项目上加密，因此这不应用作安全存储。</span><span class="sxs-lookup"><span data-stu-id="3dace-p171">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3dace-1149">参数：</span><span class="sxs-lookup"><span data-stu-id="3dace-1149">Parameters:</span></span>

|<span data-ttu-id="3dace-1150">名称</span><span class="sxs-lookup"><span data-stu-id="3dace-1150">Name</span></span>|<span data-ttu-id="3dace-1151">类型</span><span class="sxs-lookup"><span data-stu-id="3dace-1151">Type</span></span>|<span data-ttu-id="3dace-1152">属性</span><span class="sxs-lookup"><span data-stu-id="3dace-1152">Attributes</span></span>|<span data-ttu-id="3dace-1153">说明</span><span class="sxs-lookup"><span data-stu-id="3dace-1153">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="3dace-1154">函数</span><span class="sxs-lookup"><span data-stu-id="3dace-1154">function</span></span>||<span data-ttu-id="3dace-1155">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="3dace-1155">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="3dace-1156">自定义属性作为 `asyncResult.value` 属性中的 [`CustomProperties`](/javascript/api/outlook_1_7/office.customproperties) 对象提供。</span><span class="sxs-lookup"><span data-stu-id="3dace-1156">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_7/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="3dace-1157">此对象可用于获取、 设置，和从项目中删除自定义属性并将更改保存到自定义属性设置回服务器。</span><span class="sxs-lookup"><span data-stu-id="3dace-1157">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="3dace-1158">对象</span><span class="sxs-lookup"><span data-stu-id="3dace-1158">Object</span></span>|<span data-ttu-id="3dace-1159">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3dace-1159">&lt;optional&gt;</span></span>|<span data-ttu-id="3dace-1160">开发人员可以提供他们想要访问的回调函数中的任何对象。</span><span class="sxs-lookup"><span data-stu-id="3dace-1160">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="3dace-1161">此对象可以访问`asyncResult.asyncContext`的回调函数中的属性。</span><span class="sxs-lookup"><span data-stu-id="3dace-1161">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3dace-1162">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-1162">Requirements</span></span>

|<span data-ttu-id="3dace-1163">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-1163">Requirement</span></span>|<span data-ttu-id="3dace-1164">值</span><span class="sxs-lookup"><span data-stu-id="3dace-1164">Value</span></span>|
|---|---|
|[<span data-ttu-id="3dace-1165">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3dace-1165">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="3dace-1166">1.0</span><span class="sxs-lookup"><span data-stu-id="3dace-1166">1.0</span></span>|
|[<span data-ttu-id="3dace-1167">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3dace-1167">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="3dace-1168">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3dace-1168">ReadItem</span></span>|
|[<span data-ttu-id="3dace-1169">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3dace-1169">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="3dace-1170">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3dace-1170">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="3dace-1171">示例</span><span class="sxs-lookup"><span data-stu-id="3dace-1171">Example</span></span>

<span data-ttu-id="3dace-p174">以下代码示例显示了如何使用 `loadCustomPropertiesAsync` 方法异步加载特定于当前项目的自定义属性。该示例还显示了如何使用 `CustomProperties.saveAsync` 方法将这些属性重新保存到服务器。加载自定义属性后，该代码示例将使用 `CustomProperties.get` 方法读取自定义属性 `myProp`，使用 `CustomProperties.set` 方法写入自定义属性 `otherProp`，最后调用 `saveAsync` 方法保存这些自定义属性。</span><span class="sxs-lookup"><span data-stu-id="3dace-p174">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

```
// The initialize function is required for all add-ins.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
  // After the DOM is loaded, add-in-specific code can run.
  var item = Office.context.mailbox.item;
  item.loadCustomPropertiesAsync(customPropsCallback);
  });
}

function customPropsCallback(asyncResult) {
  var customProps = asyncResult.value;
  var myProp = customProps.get("myProp");

  customProps.set("otherProp", "value");
  customProps.saveAsync(saveCallback);
}

function saveCallback(asyncResult) {
}
```

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="3dace-1175">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="3dace-1175">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="3dace-1176">将附件从邮件或约会中删除。</span><span class="sxs-lookup"><span data-stu-id="3dace-1176">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="3dace-p175">`removeAttachmentAsync` 方法删除项目中带指定标识符的附件。最佳做法是，仅当同一个邮件应用程序在同一会话中添加了一个附件时，你才应使用该附件标识符来删除该附件。在 Outlook Web App 和适用于设备的 OWA 中，附件标识符只在同一个会话中才有效。当用户关闭应用程序，或者如果用户开始在内嵌窗体中撰写，并在随后弹出的内嵌窗体中继续在单独的窗口撰写时，会话即结束。</span><span class="sxs-lookup"><span data-stu-id="3dace-p175">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3dace-1181">参数：</span><span class="sxs-lookup"><span data-stu-id="3dace-1181">Parameters:</span></span>

|<span data-ttu-id="3dace-1182">名称</span><span class="sxs-lookup"><span data-stu-id="3dace-1182">Name</span></span>|<span data-ttu-id="3dace-1183">类型</span><span class="sxs-lookup"><span data-stu-id="3dace-1183">Type</span></span>|<span data-ttu-id="3dace-1184">属性</span><span class="sxs-lookup"><span data-stu-id="3dace-1184">Attributes</span></span>|<span data-ttu-id="3dace-1185">说明</span><span class="sxs-lookup"><span data-stu-id="3dace-1185">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="3dace-1186">字符串</span><span class="sxs-lookup"><span data-stu-id="3dace-1186">String</span></span>||<span data-ttu-id="3dace-p176">要删除的附件的标识符。字符串的最大长度为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="3dace-p176">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`|<span data-ttu-id="3dace-1189">对象</span><span class="sxs-lookup"><span data-stu-id="3dace-1189">Object</span></span>|<span data-ttu-id="3dace-1190">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3dace-1190">&lt;optional&gt;</span></span>|<span data-ttu-id="3dace-1191">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="3dace-1191">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="3dace-1192">对象</span><span class="sxs-lookup"><span data-stu-id="3dace-1192">Object</span></span>|<span data-ttu-id="3dace-1193">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3dace-1193">&lt;optional&gt;</span></span>|<span data-ttu-id="3dace-1194">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="3dace-1194">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="3dace-1195">函数</span><span class="sxs-lookup"><span data-stu-id="3dace-1195">function</span></span>|<span data-ttu-id="3dace-1196">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3dace-1196">&lt;optional&gt;</span></span>|<span data-ttu-id="3dace-1197">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="3dace-1197">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="3dace-1198">如果删除附件失败，`asyncResult.error` 属性将包含一个说明失败原因的错误代码。</span><span class="sxs-lookup"><span data-stu-id="3dace-1198">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="3dace-1199">错误</span><span class="sxs-lookup"><span data-stu-id="3dace-1199">Errors</span></span>

|<span data-ttu-id="3dace-1200">错误代码</span><span class="sxs-lookup"><span data-stu-id="3dace-1200">Error code</span></span>|<span data-ttu-id="3dace-1201">说明</span><span class="sxs-lookup"><span data-stu-id="3dace-1201">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="3dace-1202">附件标识符不存在。</span><span class="sxs-lookup"><span data-stu-id="3dace-1202">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3dace-1203">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-1203">Requirements</span></span>

|<span data-ttu-id="3dace-1204">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-1204">Requirement</span></span>|<span data-ttu-id="3dace-1205">值</span><span class="sxs-lookup"><span data-stu-id="3dace-1205">Value</span></span>|
|---|---|
|[<span data-ttu-id="3dace-1206">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3dace-1206">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="3dace-1207">1.1</span><span class="sxs-lookup"><span data-stu-id="3dace-1207">1.1</span></span>|
|[<span data-ttu-id="3dace-1208">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3dace-1208">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="3dace-1209">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="3dace-1209">ReadWriteItem</span></span>|
|[<span data-ttu-id="3dace-1210">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3dace-1210">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="3dace-1211">撰写</span><span class="sxs-lookup"><span data-stu-id="3dace-1211">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="3dace-1212">示例</span><span class="sxs-lookup"><span data-stu-id="3dace-1212">Example</span></span>

<span data-ttu-id="3dace-1213">以下代码删除包含标识符 '0' 的附件。</span><span class="sxs-lookup"><span data-stu-id="3dace-1213">The following code removes an attachment with an identifier of '0'.</span></span>

```
Office.context.mailbox.item.removeAttachmentAsync(
  '0',
  { asyncContext : null },
  function (asyncResult)
  {
    console.log(asyncResult.status);
  }
);
```

####  <a name="removehandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="3dace-1214">removeHandlerAsync (eventType，处理程序中，[选项] [回调])</span><span class="sxs-lookup"><span data-stu-id="3dace-1214">removeHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="3dace-1215">删除事件处理程序支持的事件。</span><span class="sxs-lookup"><span data-stu-id="3dace-1215">Removes an event handler for a supported event.</span></span>

<span data-ttu-id="3dace-1216">当前受支持的事件类型的`Office.EventType.AppointmentTimeChanged`， `Office.EventType.RecipientsChanged`，和`Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="3dace-1216">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="3dace-1217">参数：</span><span class="sxs-lookup"><span data-stu-id="3dace-1217">Parameters:</span></span>

| <span data-ttu-id="3dace-1218">名称</span><span class="sxs-lookup"><span data-stu-id="3dace-1218">Name</span></span> | <span data-ttu-id="3dace-1219">类型</span><span class="sxs-lookup"><span data-stu-id="3dace-1219">Type</span></span> | <span data-ttu-id="3dace-1220">属性</span><span class="sxs-lookup"><span data-stu-id="3dace-1220">Attributes</span></span> | <span data-ttu-id="3dace-1221">说明</span><span class="sxs-lookup"><span data-stu-id="3dace-1221">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="3dace-1222">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="3dace-1222">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="3dace-1223">应调用处理程序的事件。</span><span class="sxs-lookup"><span data-stu-id="3dace-1223">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="3dace-1224">函数</span><span class="sxs-lookup"><span data-stu-id="3dace-1224">Function</span></span> || <span data-ttu-id="3dace-p177">用于处理事件的函数。此函数必须接受一个参数，即对象文本。参数上的 `type` 属性将匹配传递给 `removeHandlerAsync` 的 `eventType` 参数。</span><span class="sxs-lookup"><span data-stu-id="3dace-p177">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `removeHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="3dace-1228">Object</span><span class="sxs-lookup"><span data-stu-id="3dace-1228">Object</span></span> | <span data-ttu-id="3dace-1229">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3dace-1229">&lt;optional&gt;</span></span> | <span data-ttu-id="3dace-1230">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="3dace-1230">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="3dace-1231">对象</span><span class="sxs-lookup"><span data-stu-id="3dace-1231">Object</span></span> | <span data-ttu-id="3dace-1232">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3dace-1232">&lt;optional&gt;</span></span> | <span data-ttu-id="3dace-1233">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="3dace-1233">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="3dace-1234">函数</span><span class="sxs-lookup"><span data-stu-id="3dace-1234">function</span></span>| <span data-ttu-id="3dace-1235">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3dace-1235">&lt;optional&gt;</span></span>|<span data-ttu-id="3dace-1236">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="3dace-1236">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3dace-1237">Requirements</span><span class="sxs-lookup"><span data-stu-id="3dace-1237">Requirements</span></span>

|<span data-ttu-id="3dace-1238">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-1238">Requirement</span></span>| <span data-ttu-id="3dace-1239">值</span><span class="sxs-lookup"><span data-stu-id="3dace-1239">Value</span></span>|
|---|---|
|[<span data-ttu-id="3dace-1240">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3dace-1240">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3dace-1241">1.7</span><span class="sxs-lookup"><span data-stu-id="3dace-1241">1.7</span></span> |
|[<span data-ttu-id="3dace-1242">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3dace-1242">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3dace-1243">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3dace-1243">ReadItem</span></span> |
|[<span data-ttu-id="3dace-1244">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3dace-1244">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3dace-1245">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3dace-1245">Compose or read</span></span> |

##### <a name="example"></a><span data-ttu-id="3dace-1246">示例</span><span class="sxs-lookup"><span data-stu-id="3dace-1246">Example</span></span>

```
Office.initialize = function (reason) {
  $(document).ready(function () {
    Office.context.mailbox.item.removeHandlerAsync(Office.EventType.RecurrenceChanged, loadNewItem, function (result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        // Handle error
      }
    });
  });
};

function loadNewItem(eventArgs) {
  // Load the properties of the newly selected item
  loadProps(Office.context.mailbox.item);
};
```

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="3dace-1247">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="3dace-1247">saveAsync([options], callback)</span></span>

<span data-ttu-id="3dace-1248">异步保存项目。</span><span class="sxs-lookup"><span data-stu-id="3dace-1248">Asynchronously saves an item.</span></span>

<span data-ttu-id="3dace-p178">调用时，此方法将当前邮件保存为草稿，并通过回调方法返回项目 ID。在 Outlook Web App 或 Outlook 联机模式下，该项目被保存到服务器中。在 Outlook 缓存模式下，该项目被保存到本地缓存中。</span><span class="sxs-lookup"><span data-stu-id="3dace-p178">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="3dace-1252">如果加载项调用`saveAsync`中的项目在撰写模式下才能获取`itemId`若要使用 EWS 或 REST API，请注意，缓存模式 Outlook 时，可能需要一些时间才能项目实际同步到服务器。</span><span class="sxs-lookup"><span data-stu-id="3dace-1252">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="3dace-1253">直到该项目同步，使用`itemId`将返回错误。</span><span class="sxs-lookup"><span data-stu-id="3dace-1253">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="3dace-p180">由于约会没有草稿状态，如果以撰写模式在约会中调用 `saveAsync`，则该项将被保存为用户日历中的正常约会。对于之前未保存过的新约会，则不会发送邀请。保存现有约会将向添加或删除的与会者发送更新。</span><span class="sxs-lookup"><span data-stu-id="3dace-p180">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="3dace-1257">以下客户端具有不同行为的`saveAsync`上约会的撰写模式下：</span><span class="sxs-lookup"><span data-stu-id="3dace-1257">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="3dace-1258">Mac Outlook 不支持`saveAsync`中的会议在撰写模式。</span><span class="sxs-lookup"><span data-stu-id="3dace-1258">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="3dace-1259">调用`saveAsync`上 Mac Outlook 中的会议，则将返回错误。</span><span class="sxs-lookup"><span data-stu-id="3dace-1259">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="3dace-1260">在 web 上的 outlook 始终发送的邀请或更新何时`saveAsync`上约会调用在撰写模式。</span><span class="sxs-lookup"><span data-stu-id="3dace-1260">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3dace-1261">参数：</span><span class="sxs-lookup"><span data-stu-id="3dace-1261">Parameters:</span></span>

|<span data-ttu-id="3dace-1262">名称</span><span class="sxs-lookup"><span data-stu-id="3dace-1262">Name</span></span>|<span data-ttu-id="3dace-1263">类型</span><span class="sxs-lookup"><span data-stu-id="3dace-1263">Type</span></span>|<span data-ttu-id="3dace-1264">属性</span><span class="sxs-lookup"><span data-stu-id="3dace-1264">Attributes</span></span>|<span data-ttu-id="3dace-1265">说明</span><span class="sxs-lookup"><span data-stu-id="3dace-1265">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="3dace-1266">对象</span><span class="sxs-lookup"><span data-stu-id="3dace-1266">Object</span></span>|<span data-ttu-id="3dace-1267">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3dace-1267">&lt;optional&gt;</span></span>|<span data-ttu-id="3dace-1268">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="3dace-1268">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="3dace-1269">对象</span><span class="sxs-lookup"><span data-stu-id="3dace-1269">Object</span></span>|<span data-ttu-id="3dace-1270">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3dace-1270">&lt;optional&gt;</span></span>|<span data-ttu-id="3dace-1271">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="3dace-1271">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="3dace-1272">函数</span><span class="sxs-lookup"><span data-stu-id="3dace-1272">function</span></span>||<span data-ttu-id="3dace-1273">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="3dace-1273">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="3dace-1274">中提供的项标识符是成功，`asyncResult.value`属性。</span><span class="sxs-lookup"><span data-stu-id="3dace-1274">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3dace-1275">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-1275">Requirements</span></span>

|<span data-ttu-id="3dace-1276">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-1276">Requirement</span></span>|<span data-ttu-id="3dace-1277">值</span><span class="sxs-lookup"><span data-stu-id="3dace-1277">Value</span></span>|
|---|---|
|[<span data-ttu-id="3dace-1278">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3dace-1278">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="3dace-1279">1.3</span><span class="sxs-lookup"><span data-stu-id="3dace-1279">1.3</span></span>|
|[<span data-ttu-id="3dace-1280">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3dace-1280">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="3dace-1281">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="3dace-1281">ReadWriteItem</span></span>|
|[<span data-ttu-id="3dace-1282">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3dace-1282">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="3dace-1283">撰写</span><span class="sxs-lookup"><span data-stu-id="3dace-1283">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="3dace-1284">示例</span><span class="sxs-lookup"><span data-stu-id="3dace-1284">Examples</span></span>

```
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="3dace-p182">下面是传递给回调函数的 `result` 参数的示例。`value` 属性包含的项目的项目 ID。</span><span class="sxs-lookup"><span data-stu-id="3dace-p182">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="3dace-1287">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="3dace-1287">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="3dace-1288">以异步方式将数据插入到邮件的正文或主题中。</span><span class="sxs-lookup"><span data-stu-id="3dace-1288">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="3dace-p183">`setSelectedDataAsync` 方法将指定字符串插入到项目主题或正文的光标位置，或者，如果在编辑器中已选择文本，则该方法将替换选择的文本。如果光标不在正文或主题字段中，则返回错误。插入之后，光标会位于插入内容的末尾。</span><span class="sxs-lookup"><span data-stu-id="3dace-p183">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3dace-1292">参数：</span><span class="sxs-lookup"><span data-stu-id="3dace-1292">Parameters:</span></span>

|<span data-ttu-id="3dace-1293">名称</span><span class="sxs-lookup"><span data-stu-id="3dace-1293">Name</span></span>|<span data-ttu-id="3dace-1294">类型</span><span class="sxs-lookup"><span data-stu-id="3dace-1294">Type</span></span>|<span data-ttu-id="3dace-1295">属性</span><span class="sxs-lookup"><span data-stu-id="3dace-1295">Attributes</span></span>|<span data-ttu-id="3dace-1296">说明</span><span class="sxs-lookup"><span data-stu-id="3dace-1296">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="3dace-1297">字符串</span><span class="sxs-lookup"><span data-stu-id="3dace-1297">String</span></span>||<span data-ttu-id="3dace-p184">要插入的数据。数据不得超过 1,000,000 个字符。如果传入的数据超过 1,000,000 个字符，则会引发 `ArgumentOutOfRange` 异常。</span><span class="sxs-lookup"><span data-stu-id="3dace-p184">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="3dace-1301">Object</span><span class="sxs-lookup"><span data-stu-id="3dace-1301">Object</span></span>|<span data-ttu-id="3dace-1302">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3dace-1302">&lt;optional&gt;</span></span>|<span data-ttu-id="3dace-1303">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="3dace-1303">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="3dace-1304">对象</span><span class="sxs-lookup"><span data-stu-id="3dace-1304">Object</span></span>|<span data-ttu-id="3dace-1305">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3dace-1305">&lt;optional&gt;</span></span>|<span data-ttu-id="3dace-1306">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="3dace-1306">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="3dace-1307">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="3dace-1307">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="3dace-1308">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="3dace-1308">&lt;optional&gt;</span></span>|<span data-ttu-id="3dace-p185">如果为 `text`，则在 Outlook Web App 和 Outlook 中应用当前样式。如果该字段是 HTML 编辑器，则仅插入文本数据，即使数据为 HTML。</span><span class="sxs-lookup"><span data-stu-id="3dace-p185">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="3dace-p186">如果 `html` 和该字段支持 HTML（主题不支持），则在 Outlook Web App 中应用当前样式，而在 Outlook 中应用默认样式。如果该字段是文本字段，则返回 `InvalidDataFormat` 错误。</span><span class="sxs-lookup"><span data-stu-id="3dace-p186">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="3dace-1313">如果未设置 `coercionType`，则结果取决于该字段：如果该字段是 HTML，则使用 HTML；如果该字段是文本，则使用纯文本。</span><span class="sxs-lookup"><span data-stu-id="3dace-1313">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="3dace-1314">function</span><span class="sxs-lookup"><span data-stu-id="3dace-1314">function</span></span>||<span data-ttu-id="3dace-1315">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="3dace-1315">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3dace-1316">Requirements</span><span class="sxs-lookup"><span data-stu-id="3dace-1316">Requirements</span></span>

|<span data-ttu-id="3dace-1317">要求</span><span class="sxs-lookup"><span data-stu-id="3dace-1317">Requirement</span></span>|<span data-ttu-id="3dace-1318">值</span><span class="sxs-lookup"><span data-stu-id="3dace-1318">Value</span></span>|
|---|---|
|[<span data-ttu-id="3dace-1319">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3dace-1319">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="3dace-1320">1.2</span><span class="sxs-lookup"><span data-stu-id="3dace-1320">1.2</span></span>|
|[<span data-ttu-id="3dace-1321">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3dace-1321">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="3dace-1322">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="3dace-1322">ReadWriteItem</span></span>|
|[<span data-ttu-id="3dace-1323">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3dace-1323">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="3dace-1324">撰写</span><span class="sxs-lookup"><span data-stu-id="3dace-1324">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="3dace-1325">示例</span><span class="sxs-lookup"><span data-stu-id="3dace-1325">Example</span></span>

```
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```