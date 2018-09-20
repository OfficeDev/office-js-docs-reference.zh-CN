
# <a name="item"></a><span data-ttu-id="67946-101">item</span><span class="sxs-lookup"><span data-stu-id="67946-101">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="67946-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="67946-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="67946-p101">`item` 命名空间用于访问当前选定的邮件、会议请求或约会。可以通过使用 [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) 属性确定 `item` 的类型。</span><span class="sxs-lookup"><span data-stu-id="67946-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="67946-105">Requirements</span><span class="sxs-lookup"><span data-stu-id="67946-105">Requirements</span></span>

|<span data-ttu-id="67946-106">要求</span><span class="sxs-lookup"><span data-stu-id="67946-106">Requirement</span></span>|<span data-ttu-id="67946-107">值</span><span class="sxs-lookup"><span data-stu-id="67946-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="67946-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="67946-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="67946-109">1.0</span><span class="sxs-lookup"><span data-stu-id="67946-109">1.0</span></span>|
|[<span data-ttu-id="67946-110">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="67946-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="67946-111">受限</span><span class="sxs-lookup"><span data-stu-id="67946-111">Restricted</span></span>|
|[<span data-ttu-id="67946-112">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="67946-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="67946-113">Compose 或 Read</span><span class="sxs-lookup"><span data-stu-id="67946-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="67946-114">成员和方法</span><span class="sxs-lookup"><span data-stu-id="67946-114">Members and methods</span></span>

| <span data-ttu-id="67946-115">成员</span><span class="sxs-lookup"><span data-stu-id="67946-115">Member</span></span> | <span data-ttu-id="67946-116">类型</span><span class="sxs-lookup"><span data-stu-id="67946-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="67946-117">attachments</span><span class="sxs-lookup"><span data-stu-id="67946-117">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails) | <span data-ttu-id="67946-118">成员</span><span class="sxs-lookup"><span data-stu-id="67946-118">Member</span></span> |
| [<span data-ttu-id="67946-119">bcc</span><span class="sxs-lookup"><span data-stu-id="67946-119">bcc</span></span>](#bcc-recipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="67946-120">成员</span><span class="sxs-lookup"><span data-stu-id="67946-120">Member</span></span> |
| [<span data-ttu-id="67946-121">body</span><span class="sxs-lookup"><span data-stu-id="67946-121">body</span></span>](#body-bodyjavascriptapioutlookofficebody) | <span data-ttu-id="67946-122">成员</span><span class="sxs-lookup"><span data-stu-id="67946-122">Member</span></span> |
| [<span data-ttu-id="67946-123">cc</span><span class="sxs-lookup"><span data-stu-id="67946-123">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="67946-124">成员</span><span class="sxs-lookup"><span data-stu-id="67946-124">Member</span></span> |
| [<span data-ttu-id="67946-125">conversationId</span><span class="sxs-lookup"><span data-stu-id="67946-125">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="67946-126">成员</span><span class="sxs-lookup"><span data-stu-id="67946-126">Member</span></span> |
| [<span data-ttu-id="67946-127">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="67946-127">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="67946-128">成员</span><span class="sxs-lookup"><span data-stu-id="67946-128">Member</span></span> |
| [<span data-ttu-id="67946-129">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="67946-129">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="67946-130">成员</span><span class="sxs-lookup"><span data-stu-id="67946-130">Member</span></span> |
| [<span data-ttu-id="67946-131">end</span><span class="sxs-lookup"><span data-stu-id="67946-131">end</span></span>](#end-datetimejavascriptapioutlookofficetime) | <span data-ttu-id="67946-132">成员</span><span class="sxs-lookup"><span data-stu-id="67946-132">Member</span></span> |
| [<span data-ttu-id="67946-133">from</span><span class="sxs-lookup"><span data-stu-id="67946-133">from</span></span>](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) | <span data-ttu-id="67946-134">成员</span><span class="sxs-lookup"><span data-stu-id="67946-134">Member</span></span> |
| [<span data-ttu-id="67946-135">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="67946-135">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="67946-136">成员</span><span class="sxs-lookup"><span data-stu-id="67946-136">Member</span></span> |
| [<span data-ttu-id="67946-137">itemClass</span><span class="sxs-lookup"><span data-stu-id="67946-137">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="67946-138">成员</span><span class="sxs-lookup"><span data-stu-id="67946-138">Member</span></span> |
| [<span data-ttu-id="67946-139">itemId</span><span class="sxs-lookup"><span data-stu-id="67946-139">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="67946-140">成员</span><span class="sxs-lookup"><span data-stu-id="67946-140">Member</span></span> |
| [<span data-ttu-id="67946-141">itemType</span><span class="sxs-lookup"><span data-stu-id="67946-141">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) | <span data-ttu-id="67946-142">成员</span><span class="sxs-lookup"><span data-stu-id="67946-142">Member</span></span> |
| [<span data-ttu-id="67946-143">location</span><span class="sxs-lookup"><span data-stu-id="67946-143">location</span></span>](#location-stringlocationjavascriptapioutlookofficelocation) | <span data-ttu-id="67946-144">成员</span><span class="sxs-lookup"><span data-stu-id="67946-144">Member</span></span> |
| [<span data-ttu-id="67946-145">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="67946-145">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="67946-146">成员</span><span class="sxs-lookup"><span data-stu-id="67946-146">Member</span></span> |
| [<span data-ttu-id="67946-147">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="67946-147">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages) | <span data-ttu-id="67946-148">成员</span><span class="sxs-lookup"><span data-stu-id="67946-148">Member</span></span> |
| [<span data-ttu-id="67946-149">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="67946-149">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="67946-150">成员</span><span class="sxs-lookup"><span data-stu-id="67946-150">Member</span></span> |
| [<span data-ttu-id="67946-151">organizer</span><span class="sxs-lookup"><span data-stu-id="67946-151">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer) | <span data-ttu-id="67946-152">成员</span><span class="sxs-lookup"><span data-stu-id="67946-152">Member</span></span> |
| [<span data-ttu-id="67946-153">recurrence</span><span class="sxs-lookup"><span data-stu-id="67946-153">recurrence</span></span>](#nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence) | <span data-ttu-id="67946-154">成员</span><span class="sxs-lookup"><span data-stu-id="67946-154">Member</span></span> |
| [<span data-ttu-id="67946-155">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="67946-155">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="67946-156">成员</span><span class="sxs-lookup"><span data-stu-id="67946-156">Member</span></span> |
| [<span data-ttu-id="67946-157">sender</span><span class="sxs-lookup"><span data-stu-id="67946-157">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) | <span data-ttu-id="67946-158">成员</span><span class="sxs-lookup"><span data-stu-id="67946-158">Member</span></span> |
| [<span data-ttu-id="67946-159">可能指向</span><span class="sxs-lookup"><span data-stu-id="67946-159">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="67946-160">成员</span><span class="sxs-lookup"><span data-stu-id="67946-160">Member</span></span> |
| [<span data-ttu-id="67946-161">start</span><span class="sxs-lookup"><span data-stu-id="67946-161">start</span></span>](#start-datetimejavascriptapioutlookofficetime) | <span data-ttu-id="67946-162">成员</span><span class="sxs-lookup"><span data-stu-id="67946-162">Member</span></span> |
| [<span data-ttu-id="67946-163">subject</span><span class="sxs-lookup"><span data-stu-id="67946-163">subject</span></span>](#subject-stringsubjectjavascriptapioutlookofficesubject) | <span data-ttu-id="67946-164">成员</span><span class="sxs-lookup"><span data-stu-id="67946-164">Member</span></span> |
| [<span data-ttu-id="67946-165">to</span><span class="sxs-lookup"><span data-stu-id="67946-165">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="67946-166">成员</span><span class="sxs-lookup"><span data-stu-id="67946-166">Member</span></span> |
| [<span data-ttu-id="67946-167">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="67946-167">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="67946-168">方法</span><span class="sxs-lookup"><span data-stu-id="67946-168">Method</span></span> |
| [<span data-ttu-id="67946-169">addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="67946-169">addFileAttachmentFromBase64Async</span></span>](#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) | <span data-ttu-id="67946-170">方法</span><span class="sxs-lookup"><span data-stu-id="67946-170">Method</span></span> |
| [<span data-ttu-id="67946-171">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="67946-171">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="67946-172">方法</span><span class="sxs-lookup"><span data-stu-id="67946-172">Method</span></span> |
| [<span data-ttu-id="67946-173">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="67946-173">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="67946-174">方法</span><span class="sxs-lookup"><span data-stu-id="67946-174">Method</span></span> |
| [<span data-ttu-id="67946-175">close</span><span class="sxs-lookup"><span data-stu-id="67946-175">close</span></span>](#close) | <span data-ttu-id="67946-176">方法</span><span class="sxs-lookup"><span data-stu-id="67946-176">Method</span></span> |
| [<span data-ttu-id="67946-177">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="67946-177">displayReplyAllForm</span></span>](#displayreplyallformformdata) | <span data-ttu-id="67946-178">方法</span><span class="sxs-lookup"><span data-stu-id="67946-178">Method</span></span> |
| [<span data-ttu-id="67946-179">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="67946-179">displayReplyForm</span></span>](#displayreplyformformdata) | <span data-ttu-id="67946-180">方法</span><span class="sxs-lookup"><span data-stu-id="67946-180">Method</span></span> |
| [<span data-ttu-id="67946-181">getEntities</span><span class="sxs-lookup"><span data-stu-id="67946-181">getEntities</span></span>](#getentities--entitiesjavascriptapioutlookofficeentities) | <span data-ttu-id="67946-182">方法</span><span class="sxs-lookup"><span data-stu-id="67946-182">Method</span></span> |
| [<span data-ttu-id="67946-183">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="67946-183">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion) | <span data-ttu-id="67946-184">方法</span><span class="sxs-lookup"><span data-stu-id="67946-184">Method</span></span> |
| [<span data-ttu-id="67946-185">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="67946-185">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion) | <span data-ttu-id="67946-186">方法</span><span class="sxs-lookup"><span data-stu-id="67946-186">Method</span></span> |
| [<span data-ttu-id="67946-187">getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="67946-187">getInitializationContextAsync</span></span>](#getinitializationcontextasyncoptions-callback) | <span data-ttu-id="67946-188">方法</span><span class="sxs-lookup"><span data-stu-id="67946-188">Method</span></span> |
| [<span data-ttu-id="67946-189">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="67946-189">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="67946-190">方法</span><span class="sxs-lookup"><span data-stu-id="67946-190">Method</span></span> |
| [<span data-ttu-id="67946-191">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="67946-191">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="67946-192">方法</span><span class="sxs-lookup"><span data-stu-id="67946-192">Method</span></span> |
| [<span data-ttu-id="67946-193">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="67946-193">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="67946-194">方法</span><span class="sxs-lookup"><span data-stu-id="67946-194">Method</span></span> |
| [<span data-ttu-id="67946-195">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="67946-195">getSelectedEntities</span></span>](#getselectedentities--entitiesjavascriptapioutlookofficeentities) | <span data-ttu-id="67946-196">方法</span><span class="sxs-lookup"><span data-stu-id="67946-196">Method</span></span> |
| [<span data-ttu-id="67946-197">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="67946-197">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="67946-198">方法</span><span class="sxs-lookup"><span data-stu-id="67946-198">Method</span></span> |
| [<span data-ttu-id="67946-199">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="67946-199">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="67946-200">方法</span><span class="sxs-lookup"><span data-stu-id="67946-200">Method</span></span> |
| [<span data-ttu-id="67946-201">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="67946-201">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="67946-202">方法</span><span class="sxs-lookup"><span data-stu-id="67946-202">Method</span></span> |
| [<span data-ttu-id="67946-203">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="67946-203">removeHandlerAsync</span></span>](#removehandlerasynceventtype-handler-options-callback) | <span data-ttu-id="67946-204">方法</span><span class="sxs-lookup"><span data-stu-id="67946-204">Method</span></span> |
| [<span data-ttu-id="67946-205">saveAsync</span><span class="sxs-lookup"><span data-stu-id="67946-205">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="67946-206">方法</span><span class="sxs-lookup"><span data-stu-id="67946-206">Method</span></span> |
| [<span data-ttu-id="67946-207">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="67946-207">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="67946-208">方法</span><span class="sxs-lookup"><span data-stu-id="67946-208">Method</span></span> |

### <a name="example"></a><span data-ttu-id="67946-209">示例</span><span class="sxs-lookup"><span data-stu-id="67946-209">Example</span></span>

<span data-ttu-id="67946-210">以下 JavaScript 代码示例显示了如何访问 Outlook 中当前项目的 `subject` 属性。</span><span class="sxs-lookup"><span data-stu-id="67946-210">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="67946-211">成员</span><span class="sxs-lookup"><span data-stu-id="67946-211">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="67946-212">attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="67946-212">attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="67946-p102">获取项目的附件数组。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="67946-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="67946-215">某些类型的文件被阻止 Outlook 由于潜在的安全问题，并因此不返回。</span><span class="sxs-lookup"><span data-stu-id="67946-215">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="67946-216">有关详细信息，请参阅[在 Outlook 中被阻止的附件](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)。</span><span class="sxs-lookup"><span data-stu-id="67946-216">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="67946-217">类型:</span><span class="sxs-lookup"><span data-stu-id="67946-217">Type:</span></span>

*   <span data-ttu-id="67946-218">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="67946-218">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="67946-219">要求</span><span class="sxs-lookup"><span data-stu-id="67946-219">Requirements</span></span>

|<span data-ttu-id="67946-220">要求</span><span class="sxs-lookup"><span data-stu-id="67946-220">Requirement</span></span>|<span data-ttu-id="67946-221">值</span><span class="sxs-lookup"><span data-stu-id="67946-221">Value</span></span>|
|---|---|
|[<span data-ttu-id="67946-222">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="67946-222">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="67946-223">1.0</span><span class="sxs-lookup"><span data-stu-id="67946-223">1.0</span></span>|
|[<span data-ttu-id="67946-224">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="67946-224">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="67946-225">ReadItem</span><span class="sxs-lookup"><span data-stu-id="67946-225">ReadItem</span></span>|
|[<span data-ttu-id="67946-226">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="67946-226">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="67946-227">阅读</span><span class="sxs-lookup"><span data-stu-id="67946-227">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="67946-228">示例</span><span class="sxs-lookup"><span data-stu-id="67946-228">Example</span></span>

<span data-ttu-id="67946-229">以下代码使用当前项目上所有附件的详细信息构成 HTML 字符串。</span><span class="sxs-lookup"><span data-stu-id="67946-229">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="67946-230">bcc :[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="67946-230">bcc :[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="67946-231">获取一个对象，提供用于获取或更新邮件的密件抄送 （密件抄送副本） 行上的收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="67946-231">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="67946-232">仅限撰写模式。</span><span class="sxs-lookup"><span data-stu-id="67946-232">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="67946-233">类型：</span><span class="sxs-lookup"><span data-stu-id="67946-233">Type:</span></span>

*   [<span data-ttu-id="67946-234">收件人</span><span class="sxs-lookup"><span data-stu-id="67946-234">Recipients</span></span>](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="67946-235">要求</span><span class="sxs-lookup"><span data-stu-id="67946-235">Requirements</span></span>

|<span data-ttu-id="67946-236">要求</span><span class="sxs-lookup"><span data-stu-id="67946-236">Requirement</span></span>|<span data-ttu-id="67946-237">值</span><span class="sxs-lookup"><span data-stu-id="67946-237">Value</span></span>|
|---|---|
|[<span data-ttu-id="67946-238">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="67946-238">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="67946-239">1.1</span><span class="sxs-lookup"><span data-stu-id="67946-239">1.1</span></span>|
|[<span data-ttu-id="67946-240">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="67946-240">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="67946-241">ReadItem</span><span class="sxs-lookup"><span data-stu-id="67946-241">ReadItem</span></span>|
|[<span data-ttu-id="67946-242">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="67946-242">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="67946-243">撰写</span><span class="sxs-lookup"><span data-stu-id="67946-243">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="67946-244">示例</span><span class="sxs-lookup"><span data-stu-id="67946-244">Example</span></span>

```
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlookofficebody"></a><span data-ttu-id="67946-245">body :[Body](/javascript/api/outlook/office.body)</span><span class="sxs-lookup"><span data-stu-id="67946-245">body :[Body](/javascript/api/outlook/office.body)</span></span>

<span data-ttu-id="67946-246">获取一个提供用于处理项目正文的方法的对象。</span><span class="sxs-lookup"><span data-stu-id="67946-246">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="67946-247">类型：</span><span class="sxs-lookup"><span data-stu-id="67946-247">Type:</span></span>

*   [<span data-ttu-id="67946-248">Body</span><span class="sxs-lookup"><span data-stu-id="67946-248">Body</span></span>](/javascript/api/outlook/office.body)

##### <a name="requirements"></a><span data-ttu-id="67946-249">要求</span><span class="sxs-lookup"><span data-stu-id="67946-249">Requirements</span></span>

|<span data-ttu-id="67946-250">要求</span><span class="sxs-lookup"><span data-stu-id="67946-250">Requirement</span></span>|<span data-ttu-id="67946-251">值</span><span class="sxs-lookup"><span data-stu-id="67946-251">Value</span></span>|
|---|---|
|[<span data-ttu-id="67946-252">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="67946-252">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="67946-253">1.1</span><span class="sxs-lookup"><span data-stu-id="67946-253">1.1</span></span>|
|[<span data-ttu-id="67946-254">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="67946-254">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="67946-255">ReadItem</span><span class="sxs-lookup"><span data-stu-id="67946-255">ReadItem</span></span>|
|[<span data-ttu-id="67946-256">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="67946-256">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="67946-257">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="67946-257">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="67946-258">抄送： 数组。 <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[收件人](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="67946-258">cc :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="67946-259">提供对邮件的抄送 (cc) 收件人访问。</span><span class="sxs-lookup"><span data-stu-id="67946-259">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="67946-260">对象的类型和访问级别取决于当前项的模式。</span><span class="sxs-lookup"><span data-stu-id="67946-260">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="67946-261">阅读模式</span><span class="sxs-lookup"><span data-stu-id="67946-261">Read mode</span></span>

<span data-ttu-id="67946-p106">`cc` 属性返回包含邮件的**抄送**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="67946-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="67946-264">撰写模式</span><span class="sxs-lookup"><span data-stu-id="67946-264">Compose mode</span></span>

<span data-ttu-id="67946-265">`cc`属性返回`Recipients`提供方法用于获取或更新的邮件**Cc**行上的收件人对象。</span><span class="sxs-lookup"><span data-stu-id="67946-265">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="67946-266">类型：</span><span class="sxs-lookup"><span data-stu-id="67946-266">Type:</span></span>

*   <span data-ttu-id="67946-267">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="67946-267">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="67946-268">要求</span><span class="sxs-lookup"><span data-stu-id="67946-268">Requirements</span></span>

|<span data-ttu-id="67946-269">要求</span><span class="sxs-lookup"><span data-stu-id="67946-269">Requirement</span></span>|<span data-ttu-id="67946-270">值</span><span class="sxs-lookup"><span data-stu-id="67946-270">Value</span></span>|
|---|---|
|[<span data-ttu-id="67946-271">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="67946-271">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="67946-272">1.0</span><span class="sxs-lookup"><span data-stu-id="67946-272">1.0</span></span>|
|[<span data-ttu-id="67946-273">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="67946-273">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="67946-274">ReadItem</span><span class="sxs-lookup"><span data-stu-id="67946-274">ReadItem</span></span>|
|[<span data-ttu-id="67946-275">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="67946-275">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="67946-276">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="67946-276">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="67946-277">示例</span><span class="sxs-lookup"><span data-stu-id="67946-277">Example</span></span>

```
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="67946-278">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="67946-278">(nullable) conversationId :String</span></span>

<span data-ttu-id="67946-279">获取包含特定消息的电子邮件会话的标识符。</span><span class="sxs-lookup"><span data-stu-id="67946-279">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="67946-p107">如果在阅读窗体或撰写窗体的回复中激活邮件应用程序，则此属性可以获得一个整数值。如果用户随后更改了回复邮件的主题（若发送回复），则该邮件的对话 ID 将改变且之前获取的值将不适用。</span><span class="sxs-lookup"><span data-stu-id="67946-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="67946-p108">对于撰写窗体的新项目，此属性获得一个 null 值。如果用户设置一个主题并保存该项目，`conversationId` 属性将返回一个值。</span><span class="sxs-lookup"><span data-stu-id="67946-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="67946-284">类型:</span><span class="sxs-lookup"><span data-stu-id="67946-284">Type:</span></span>

*   <span data-ttu-id="67946-285">String</span><span class="sxs-lookup"><span data-stu-id="67946-285">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="67946-286">要求</span><span class="sxs-lookup"><span data-stu-id="67946-286">Requirements</span></span>

|<span data-ttu-id="67946-287">要求</span><span class="sxs-lookup"><span data-stu-id="67946-287">Requirement</span></span>|<span data-ttu-id="67946-288">值</span><span class="sxs-lookup"><span data-stu-id="67946-288">Value</span></span>|
|---|---|
|[<span data-ttu-id="67946-289">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="67946-289">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="67946-290">1.0</span><span class="sxs-lookup"><span data-stu-id="67946-290">1.0</span></span>|
|[<span data-ttu-id="67946-291">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="67946-291">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="67946-292">ReadItem</span><span class="sxs-lookup"><span data-stu-id="67946-292">ReadItem</span></span>|
|[<span data-ttu-id="67946-293">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="67946-293">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="67946-294">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="67946-294">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="67946-295">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="67946-295">dateTimeCreated :Date</span></span>

<span data-ttu-id="67946-p109">获取项目创建的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="67946-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="67946-298">类型：</span><span class="sxs-lookup"><span data-stu-id="67946-298">Type:</span></span>

*   <span data-ttu-id="67946-299">日期</span><span class="sxs-lookup"><span data-stu-id="67946-299">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="67946-300">要求</span><span class="sxs-lookup"><span data-stu-id="67946-300">Requirements</span></span>

|<span data-ttu-id="67946-301">要求</span><span class="sxs-lookup"><span data-stu-id="67946-301">Requirement</span></span>|<span data-ttu-id="67946-302">值</span><span class="sxs-lookup"><span data-stu-id="67946-302">Value</span></span>|
|---|---|
|[<span data-ttu-id="67946-303">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="67946-303">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="67946-304">1.0</span><span class="sxs-lookup"><span data-stu-id="67946-304">1.0</span></span>|
|[<span data-ttu-id="67946-305">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="67946-305">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="67946-306">ReadItem</span><span class="sxs-lookup"><span data-stu-id="67946-306">ReadItem</span></span>|
|[<span data-ttu-id="67946-307">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="67946-307">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="67946-308">阅读</span><span class="sxs-lookup"><span data-stu-id="67946-308">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="67946-309">示例</span><span class="sxs-lookup"><span data-stu-id="67946-309">Example</span></span>

```
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="67946-310">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="67946-310">dateTimeModified :Date</span></span>

<span data-ttu-id="67946-p110">获取项目最近一次修改的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="67946-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="67946-313">此成员不支持适用于 iOS 的 Outlook 或 Outlook for Android。</span><span class="sxs-lookup"><span data-stu-id="67946-313">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="67946-314">类型:</span><span class="sxs-lookup"><span data-stu-id="67946-314">Type:</span></span>

*   <span data-ttu-id="67946-315">日期</span><span class="sxs-lookup"><span data-stu-id="67946-315">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="67946-316">要求</span><span class="sxs-lookup"><span data-stu-id="67946-316">Requirements</span></span>

|<span data-ttu-id="67946-317">要求</span><span class="sxs-lookup"><span data-stu-id="67946-317">Requirement</span></span>|<span data-ttu-id="67946-318">值</span><span class="sxs-lookup"><span data-stu-id="67946-318">Value</span></span>|
|---|---|
|[<span data-ttu-id="67946-319">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="67946-319">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="67946-320">1.0</span><span class="sxs-lookup"><span data-stu-id="67946-320">1.0</span></span>|
|[<span data-ttu-id="67946-321">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="67946-321">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="67946-322">ReadItem</span><span class="sxs-lookup"><span data-stu-id="67946-322">ReadItem</span></span>|
|[<span data-ttu-id="67946-323">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="67946-323">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="67946-324">阅读</span><span class="sxs-lookup"><span data-stu-id="67946-324">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="67946-325">示例</span><span class="sxs-lookup"><span data-stu-id="67946-325">Example</span></span>

```
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="67946-326">end :Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="67946-326">end :Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="67946-327">获取或设置约会结束的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="67946-327">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="67946-p111">将 `end` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) 方法将 end 属性值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="67946-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="67946-330">阅读模式</span><span class="sxs-lookup"><span data-stu-id="67946-330">Read mode</span></span>

<span data-ttu-id="67946-331">`end` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="67946-331">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="67946-332">撰写模式</span><span class="sxs-lookup"><span data-stu-id="67946-332">Compose mode</span></span>

<span data-ttu-id="67946-333">`end` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="67946-333">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="67946-334">使用 [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) 方法设置结束时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="67946-334">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="67946-335">类型：</span><span class="sxs-lookup"><span data-stu-id="67946-335">Type:</span></span>

*   <span data-ttu-id="67946-336">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="67946-336">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="67946-337">要求</span><span class="sxs-lookup"><span data-stu-id="67946-337">Requirements</span></span>

|<span data-ttu-id="67946-338">要求</span><span class="sxs-lookup"><span data-stu-id="67946-338">Requirement</span></span>|<span data-ttu-id="67946-339">值</span><span class="sxs-lookup"><span data-stu-id="67946-339">Value</span></span>|
|---|---|
|[<span data-ttu-id="67946-340">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="67946-340">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="67946-341">1.0</span><span class="sxs-lookup"><span data-stu-id="67946-341">1.0</span></span>|
|[<span data-ttu-id="67946-342">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="67946-342">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="67946-343">ReadItem</span><span class="sxs-lookup"><span data-stu-id="67946-343">ReadItem</span></span>|
|[<span data-ttu-id="67946-344">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="67946-344">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="67946-345">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="67946-345">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="67946-346">示例</span><span class="sxs-lookup"><span data-stu-id="67946-346">Example</span></span>

<span data-ttu-id="67946-347">以下示例通过使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) 方法，设置撰写模式下约会的结束时间。</span><span class="sxs-lookup"><span data-stu-id="67946-347">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom"></a><span data-ttu-id="67946-348">从：[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[从](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="67946-348">from :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span></span>

<span data-ttu-id="67946-349">获取邮件发件人的电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="67946-349">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="67946-p112">`from` 和 [`sender`](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="67946-p112">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="67946-352">`recipientType`属性`EmailAddressDetails`对象在`from`属性是`undefined`。</span><span class="sxs-lookup"><span data-stu-id="67946-352">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="67946-353">阅读模式</span><span class="sxs-lookup"><span data-stu-id="67946-353">Read mode</span></span>

<span data-ttu-id="67946-354">`from`属性返回`EmailAddressDetails`对象。</span><span class="sxs-lookup"><span data-stu-id="67946-354">The `from` property returns an `EmailAddressDetails` object.</span></span>

```
var subject = Office.context.mailbox.item.from;
```

##### <a name="compose-mode"></a><span data-ttu-id="67946-355">撰写模式</span><span class="sxs-lookup"><span data-stu-id="67946-355">Compose mode</span></span>

<span data-ttu-id="67946-356">`from`属性返回`From`对象，它提供一个方法来获取值。</span><span class="sxs-lookup"><span data-stu-id="67946-356">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="67946-357">类型:</span><span class="sxs-lookup"><span data-stu-id="67946-357">Type:</span></span>

*   <span data-ttu-id="67946-358">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [从](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="67946-358">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="67946-359">要求</span><span class="sxs-lookup"><span data-stu-id="67946-359">Requirements</span></span>

|<span data-ttu-id="67946-360">要求</span><span class="sxs-lookup"><span data-stu-id="67946-360">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="67946-361">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="67946-361">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="67946-362">1.0</span><span class="sxs-lookup"><span data-stu-id="67946-362">1.0</span></span>|<span data-ttu-id="67946-363">预览</span><span class="sxs-lookup"><span data-stu-id="67946-363">Preview</span></span>|
|[<span data-ttu-id="67946-364">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="67946-364">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="67946-365">ReadItem</span><span class="sxs-lookup"><span data-stu-id="67946-365">ReadItem</span></span>|<span data-ttu-id="67946-366">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="67946-366">ReadWriteItem</span></span>|
|[<span data-ttu-id="67946-367">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="67946-367">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="67946-368">Read</span><span class="sxs-lookup"><span data-stu-id="67946-368">Read</span></span>|<span data-ttu-id="67946-369">撰写</span><span class="sxs-lookup"><span data-stu-id="67946-369">Compose</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="67946-370">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="67946-370">internetMessageId :String</span></span>

<span data-ttu-id="67946-p113">获取电子邮件的 Internet 消息标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="67946-p113">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="67946-373">类型：</span><span class="sxs-lookup"><span data-stu-id="67946-373">Type:</span></span>

*   <span data-ttu-id="67946-374">String</span><span class="sxs-lookup"><span data-stu-id="67946-374">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="67946-375">要求</span><span class="sxs-lookup"><span data-stu-id="67946-375">Requirements</span></span>

|<span data-ttu-id="67946-376">要求</span><span class="sxs-lookup"><span data-stu-id="67946-376">Requirement</span></span>|<span data-ttu-id="67946-377">值</span><span class="sxs-lookup"><span data-stu-id="67946-377">Value</span></span>|
|---|---|
|[<span data-ttu-id="67946-378">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="67946-378">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="67946-379">1.0</span><span class="sxs-lookup"><span data-stu-id="67946-379">1.0</span></span>|
|[<span data-ttu-id="67946-380">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="67946-380">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="67946-381">ReadItem</span><span class="sxs-lookup"><span data-stu-id="67946-381">ReadItem</span></span>|
|[<span data-ttu-id="67946-382">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="67946-382">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="67946-383">阅读</span><span class="sxs-lookup"><span data-stu-id="67946-383">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="67946-384">示例</span><span class="sxs-lookup"><span data-stu-id="67946-384">Example</span></span>

```
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="67946-385">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="67946-385">itemClass :String</span></span>

<span data-ttu-id="67946-p114">获取选定项目的 Exchange Web 服务项目类。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="67946-p114">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="67946-p115">`itemClass` 属性指定所选项目的邮件类别。以下是邮件或约会项目的默认邮件类别。</span><span class="sxs-lookup"><span data-stu-id="67946-p115">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="67946-390">类型</span><span class="sxs-lookup"><span data-stu-id="67946-390">Type</span></span>|<span data-ttu-id="67946-391">说明</span><span class="sxs-lookup"><span data-stu-id="67946-391">Description</span></span>|<span data-ttu-id="67946-392">项目类</span><span class="sxs-lookup"><span data-stu-id="67946-392">item class</span></span>|
|---|---|---|
|<span data-ttu-id="67946-393">约会项目</span><span class="sxs-lookup"><span data-stu-id="67946-393">Appointment items</span></span>|<span data-ttu-id="67946-394">这些是项目类 `IPM.Appointment` 或 `IPM.Appointment.Occurence` 的日历项目。</span><span class="sxs-lookup"><span data-stu-id="67946-394">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurence`|
|<span data-ttu-id="67946-395">邮件项目</span><span class="sxs-lookup"><span data-stu-id="67946-395">Message items</span></span>|<span data-ttu-id="67946-396">这些项目包括具有默认邮件类别 `IPM.Note` 的电子邮件，以及使用 `IPM.Schedule.Meeting` 作为基础邮件类别的会议请求、响应和取消。</span><span class="sxs-lookup"><span data-stu-id="67946-396">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="67946-397">你可以创建用于扩展默认邮件类别的自定义邮件类别，例如，自定义约会邮件类别 `IPM.Appointment.Contoso`。</span><span class="sxs-lookup"><span data-stu-id="67946-397">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="67946-398">类型：</span><span class="sxs-lookup"><span data-stu-id="67946-398">Type:</span></span>

*   <span data-ttu-id="67946-399">String</span><span class="sxs-lookup"><span data-stu-id="67946-399">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="67946-400">要求</span><span class="sxs-lookup"><span data-stu-id="67946-400">Requirements</span></span>

|<span data-ttu-id="67946-401">要求</span><span class="sxs-lookup"><span data-stu-id="67946-401">Requirement</span></span>|<span data-ttu-id="67946-402">值</span><span class="sxs-lookup"><span data-stu-id="67946-402">Value</span></span>|
|---|---|
|[<span data-ttu-id="67946-403">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="67946-403">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="67946-404">1.0</span><span class="sxs-lookup"><span data-stu-id="67946-404">1.0</span></span>|
|[<span data-ttu-id="67946-405">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="67946-405">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="67946-406">ReadItem</span><span class="sxs-lookup"><span data-stu-id="67946-406">ReadItem</span></span>|
|[<span data-ttu-id="67946-407">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="67946-407">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="67946-408">阅读</span><span class="sxs-lookup"><span data-stu-id="67946-408">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="67946-409">示例</span><span class="sxs-lookup"><span data-stu-id="67946-409">Example</span></span>

```
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="67946-410">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="67946-410">(nullable) itemId :String</span></span>

<span data-ttu-id="67946-p116">获取当前项目的 Exchange Web 服务项目标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="67946-p116">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="67946-413">`itemId` 属性返回的标识符与 Exchange Web 服务项目标识符相同。</span><span class="sxs-lookup"><span data-stu-id="67946-413">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="67946-414">`itemId`属性不是与 Outlook 条目 ID 或使用 Outlook REST API 的 ID。</span><span class="sxs-lookup"><span data-stu-id="67946-414">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="67946-415">使用此值的 REST API 调用之前，应将其转换使用[Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string)。</span><span class="sxs-lookup"><span data-stu-id="67946-415">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="67946-416">有关详细信息，请参阅[使用 Outlook REST Api 从 Outlook 外接程序](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id)。</span><span class="sxs-lookup"><span data-stu-id="67946-416">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="67946-p118">`itemId` 属性在撰写模式下不可用。如果需要项目标识符，[`saveAsync`](#saveasyncoptions-callback) 方法可用于将项目保存到存储，这将在回调函数的 [`AsyncResult.value`](/javascript/api/office/office.asyncresult) 参数中返回项目标识符。</span><span class="sxs-lookup"><span data-stu-id="67946-p118">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="67946-419">类型:</span><span class="sxs-lookup"><span data-stu-id="67946-419">Type:</span></span>

*   <span data-ttu-id="67946-420">String</span><span class="sxs-lookup"><span data-stu-id="67946-420">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="67946-421">要求</span><span class="sxs-lookup"><span data-stu-id="67946-421">Requirements</span></span>

|<span data-ttu-id="67946-422">要求</span><span class="sxs-lookup"><span data-stu-id="67946-422">Requirement</span></span>|<span data-ttu-id="67946-423">值</span><span class="sxs-lookup"><span data-stu-id="67946-423">Value</span></span>|
|---|---|
|[<span data-ttu-id="67946-424">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="67946-424">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="67946-425">1.0</span><span class="sxs-lookup"><span data-stu-id="67946-425">1.0</span></span>|
|[<span data-ttu-id="67946-426">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="67946-426">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="67946-427">ReadItem</span><span class="sxs-lookup"><span data-stu-id="67946-427">ReadItem</span></span>|
|[<span data-ttu-id="67946-428">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="67946-428">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="67946-429">阅读</span><span class="sxs-lookup"><span data-stu-id="67946-429">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="67946-430">示例</span><span class="sxs-lookup"><span data-stu-id="67946-430">Example</span></span>

<span data-ttu-id="67946-p119">以下代码检查项目标识符是否存在。如果 `itemId` 属性返回 `null` 或 `undefined`，则将项目保存到存储，并从异步结果中获取项目标识符。</span><span class="sxs-lookup"><span data-stu-id="67946-p119">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype"></a><span data-ttu-id="67946-433">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="67946-433">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="67946-434">获取实例表示的项的类型。</span><span class="sxs-lookup"><span data-stu-id="67946-434">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="67946-435">`itemType` 属性返回其中一个 `ItemType` 枚举值，指示 `item` 对象实例是邮件还是约会。</span><span class="sxs-lookup"><span data-stu-id="67946-435">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="67946-436">类型：</span><span class="sxs-lookup"><span data-stu-id="67946-436">Type:</span></span>

*   [<span data-ttu-id="67946-437">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="67946-437">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="67946-438">要求</span><span class="sxs-lookup"><span data-stu-id="67946-438">Requirements</span></span>

|<span data-ttu-id="67946-439">要求</span><span class="sxs-lookup"><span data-stu-id="67946-439">Requirement</span></span>|<span data-ttu-id="67946-440">值</span><span class="sxs-lookup"><span data-stu-id="67946-440">Value</span></span>|
|---|---|
|[<span data-ttu-id="67946-441">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="67946-441">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="67946-442">1.0</span><span class="sxs-lookup"><span data-stu-id="67946-442">1.0</span></span>|
|[<span data-ttu-id="67946-443">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="67946-443">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="67946-444">ReadItem</span><span class="sxs-lookup"><span data-stu-id="67946-444">ReadItem</span></span>|
|[<span data-ttu-id="67946-445">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="67946-445">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="67946-446">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="67946-446">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="67946-447">示例</span><span class="sxs-lookup"><span data-stu-id="67946-447">Example</span></span>

```
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlookofficelocation"></a><span data-ttu-id="67946-448">location :String|[Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="67946-448">location :String|[Location](/javascript/api/outlook/office.location)</span></span>

<span data-ttu-id="67946-449">获取或设置约会的位置。</span><span class="sxs-lookup"><span data-stu-id="67946-449">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="67946-450">阅读模式</span><span class="sxs-lookup"><span data-stu-id="67946-450">Read mode</span></span>

<span data-ttu-id="67946-451">`location` 属性返回一个包含约会位置的字符串。</span><span class="sxs-lookup"><span data-stu-id="67946-451">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="67946-452">撰写模式</span><span class="sxs-lookup"><span data-stu-id="67946-452">Compose mode</span></span>

<span data-ttu-id="67946-453">`location` 属性返回一个 `Location` 对象，该对象提供用于获取和设置约会位置的方法。</span><span class="sxs-lookup"><span data-stu-id="67946-453">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="67946-454">类型：</span><span class="sxs-lookup"><span data-stu-id="67946-454">Type:</span></span>

*   <span data-ttu-id="67946-455">String | [Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="67946-455">String | [Location](/javascript/api/outlook/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="67946-456">要求</span><span class="sxs-lookup"><span data-stu-id="67946-456">Requirements</span></span>

|<span data-ttu-id="67946-457">要求</span><span class="sxs-lookup"><span data-stu-id="67946-457">Requirement</span></span>|<span data-ttu-id="67946-458">值</span><span class="sxs-lookup"><span data-stu-id="67946-458">Value</span></span>|
|---|---|
|[<span data-ttu-id="67946-459">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="67946-459">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="67946-460">1.0</span><span class="sxs-lookup"><span data-stu-id="67946-460">1.0</span></span>|
|[<span data-ttu-id="67946-461">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="67946-461">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="67946-462">ReadItem</span><span class="sxs-lookup"><span data-stu-id="67946-462">ReadItem</span></span>|
|[<span data-ttu-id="67946-463">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="67946-463">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="67946-464">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="67946-464">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="67946-465">示例</span><span class="sxs-lookup"><span data-stu-id="67946-465">Example</span></span>

```
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="67946-466">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="67946-466">normalizedSubject :String</span></span>

<span data-ttu-id="67946-p120">获取删除了所有前缀（包括 `RE:` 和 `FWD:`）的项目主题。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="67946-p120">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="67946-p121">normalizedSubject 属性获取包含由电子邮件程序添加的任何标准前缀（如 `RE:` 和 `FW:`）的项目主题。若要获取包含完整前缀的项目主题，请使用 [`subject`](#subject-stringsubjectjavascriptapioutlookofficesubject) 属性。</span><span class="sxs-lookup"><span data-stu-id="67946-p121">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlookofficesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="67946-471">类型:</span><span class="sxs-lookup"><span data-stu-id="67946-471">Type:</span></span>

*   <span data-ttu-id="67946-472">String</span><span class="sxs-lookup"><span data-stu-id="67946-472">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="67946-473">要求</span><span class="sxs-lookup"><span data-stu-id="67946-473">Requirements</span></span>

|<span data-ttu-id="67946-474">要求</span><span class="sxs-lookup"><span data-stu-id="67946-474">Requirement</span></span>|<span data-ttu-id="67946-475">值</span><span class="sxs-lookup"><span data-stu-id="67946-475">Value</span></span>|
|---|---|
|[<span data-ttu-id="67946-476">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="67946-476">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="67946-477">1.0</span><span class="sxs-lookup"><span data-stu-id="67946-477">1.0</span></span>|
|[<span data-ttu-id="67946-478">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="67946-478">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="67946-479">ReadItem</span><span class="sxs-lookup"><span data-stu-id="67946-479">ReadItem</span></span>|
|[<span data-ttu-id="67946-480">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="67946-480">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="67946-481">阅读</span><span class="sxs-lookup"><span data-stu-id="67946-481">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="67946-482">示例</span><span class="sxs-lookup"><span data-stu-id="67946-482">Example</span></span>

```
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages"></a><span data-ttu-id="67946-483">notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="67946-483">notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span></span>

<span data-ttu-id="67946-484">获取项目的通知邮件。</span><span class="sxs-lookup"><span data-stu-id="67946-484">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="67946-485">类型：</span><span class="sxs-lookup"><span data-stu-id="67946-485">Type:</span></span>

*   [<span data-ttu-id="67946-486">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="67946-486">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="67946-487">要求</span><span class="sxs-lookup"><span data-stu-id="67946-487">Requirements</span></span>

|<span data-ttu-id="67946-488">要求</span><span class="sxs-lookup"><span data-stu-id="67946-488">Requirement</span></span>|<span data-ttu-id="67946-489">值</span><span class="sxs-lookup"><span data-stu-id="67946-489">Value</span></span>|
|---|---|
|[<span data-ttu-id="67946-490">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="67946-490">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="67946-491">1.3</span><span class="sxs-lookup"><span data-stu-id="67946-491">1.3</span></span>|
|[<span data-ttu-id="67946-492">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="67946-492">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="67946-493">ReadItem</span><span class="sxs-lookup"><span data-stu-id="67946-493">ReadItem</span></span>|
|[<span data-ttu-id="67946-494">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="67946-494">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="67946-495">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="67946-495">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="67946-496">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="67946-496">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="67946-497">提供对事件的可选与会者的访问。</span><span class="sxs-lookup"><span data-stu-id="67946-497">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="67946-498">对象的类型和访问级别取决于当前项的模式。</span><span class="sxs-lookup"><span data-stu-id="67946-498">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="67946-499">阅读模式</span><span class="sxs-lookup"><span data-stu-id="67946-499">Read mode</span></span>

<span data-ttu-id="67946-500">`optionalAttendees` 属性返回一个数组，其中包含每个可选与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="67946-500">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="67946-501">撰写模式</span><span class="sxs-lookup"><span data-stu-id="67946-501">Compose mode</span></span>

<span data-ttu-id="67946-502">`optionalAttendees`属性返回`Recipients`对象，它提供用于获取或更新会议的可选与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="67946-502">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="67946-503">类型：</span><span class="sxs-lookup"><span data-stu-id="67946-503">Type:</span></span>

*   <span data-ttu-id="67946-504">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="67946-504">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="67946-505">要求</span><span class="sxs-lookup"><span data-stu-id="67946-505">Requirements</span></span>

|<span data-ttu-id="67946-506">要求</span><span class="sxs-lookup"><span data-stu-id="67946-506">Requirement</span></span>|<span data-ttu-id="67946-507">值</span><span class="sxs-lookup"><span data-stu-id="67946-507">Value</span></span>|
|---|---|
|[<span data-ttu-id="67946-508">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="67946-508">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="67946-509">1.0</span><span class="sxs-lookup"><span data-stu-id="67946-509">1.0</span></span>|
|[<span data-ttu-id="67946-510">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="67946-510">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="67946-511">ReadItem</span><span class="sxs-lookup"><span data-stu-id="67946-511">ReadItem</span></span>|
|[<span data-ttu-id="67946-512">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="67946-512">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="67946-513">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="67946-513">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="67946-514">示例</span><span class="sxs-lookup"><span data-stu-id="67946-514">Example</span></span>

```
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer"></a><span data-ttu-id="67946-515">组织者：[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[组织者](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="67946-515">organizer :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span></span>

<span data-ttu-id="67946-516">获取指定会议组织者的电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="67946-516">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="67946-517">阅读模式</span><span class="sxs-lookup"><span data-stu-id="67946-517">Read mode</span></span>

<span data-ttu-id="67946-518">`organizer`属性返回一个代表会议组织者的[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)对象。</span><span class="sxs-lookup"><span data-stu-id="67946-518">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="67946-519">撰写模式</span><span class="sxs-lookup"><span data-stu-id="67946-519">Compose mode</span></span>

<span data-ttu-id="67946-520">`organizer`属性返回一个[管理器](/javascript/api/outlook/office.organizer)对象，提供要获取组织者值的方法。</span><span class="sxs-lookup"><span data-stu-id="67946-520">The `organizer` property returns an [Organizer](/javascript/api/outlook/office.organizer) object that provides a method to get the organizer value.</span></span>

##### <a name="type"></a><span data-ttu-id="67946-521">类型:</span><span class="sxs-lookup"><span data-stu-id="67946-521">Type:</span></span>

*   <span data-ttu-id="67946-522">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [组织者](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="67946-522">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="67946-523">要求</span><span class="sxs-lookup"><span data-stu-id="67946-523">Requirements</span></span>

|<span data-ttu-id="67946-524">要求</span><span class="sxs-lookup"><span data-stu-id="67946-524">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="67946-525">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="67946-525">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="67946-526">1.0</span><span class="sxs-lookup"><span data-stu-id="67946-526">1.0</span></span>|<span data-ttu-id="67946-527">预览</span><span class="sxs-lookup"><span data-stu-id="67946-527">Preview</span></span>|
|[<span data-ttu-id="67946-528">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="67946-528">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="67946-529">ReadItem</span><span class="sxs-lookup"><span data-stu-id="67946-529">ReadItem</span></span>|<span data-ttu-id="67946-530">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="67946-530">ReadWriteItem</span></span>|
|[<span data-ttu-id="67946-531">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="67946-531">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="67946-532">Read</span><span class="sxs-lookup"><span data-stu-id="67946-532">Read</span></span>|<span data-ttu-id="67946-533">撰写</span><span class="sxs-lookup"><span data-stu-id="67946-533">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="67946-534">示例</span><span class="sxs-lookup"><span data-stu-id="67946-534">Example</span></span>

```
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence"></a><span data-ttu-id="67946-535">(可以为 null) 的定期：[定期](/javascript/api/outlook/office.recurrence)</span><span class="sxs-lookup"><span data-stu-id="67946-535">(nullable) recurrence :[Recurrence](/javascript/api/outlook/office.recurrence)</span></span>

<span data-ttu-id="67946-536">获取或设置约会的定期模式。</span><span class="sxs-lookup"><span data-stu-id="67946-536">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="67946-537">获取会议请求的定期模式。</span><span class="sxs-lookup"><span data-stu-id="67946-537">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="67946-538">阅读和撰写的约会项目的模式。</span><span class="sxs-lookup"><span data-stu-id="67946-538">Read and compose modes for appointment items.</span></span> <span data-ttu-id="67946-539">读取模式的会议请求项目。</span><span class="sxs-lookup"><span data-stu-id="67946-539">Read mode for meeting request items.</span></span>

<span data-ttu-id="67946-540">`recurrence`属性返回的定期约会或会议请求的[定期](/javascript/api/outlook/office.recurrence)对象，如果项是一系列或一系列中的实例。</span><span class="sxs-lookup"><span data-stu-id="67946-540">The `recurrence` property returns a [recurrence](/javascript/api/outlook/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="67946-541">`null`将返回单个约会和会议请求的单个约会。</span><span class="sxs-lookup"><span data-stu-id="67946-541">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="67946-542">`undefined`将返回不是会议请求的消息。</span><span class="sxs-lookup"><span data-stu-id="67946-542">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="67946-543">注意： 会议请求具有`itemClass`IPM 的值。Schedule.Meeting.Request。</span><span class="sxs-lookup"><span data-stu-id="67946-543">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="67946-544">注意： 如果定期对象是`null`，这表明该对象是一个约会或会议请求的一个约会，不属于一系列。</span><span class="sxs-lookup"><span data-stu-id="67946-544">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="type"></a><span data-ttu-id="67946-545">类型:</span><span class="sxs-lookup"><span data-stu-id="67946-545">Type:</span></span>

* [<span data-ttu-id="67946-546">定期</span><span class="sxs-lookup"><span data-stu-id="67946-546">Recurrence</span></span>](/javascript/api/outlook/office.recurrence)

|<span data-ttu-id="67946-547">要求</span><span class="sxs-lookup"><span data-stu-id="67946-547">Requirement</span></span>|<span data-ttu-id="67946-548">值</span><span class="sxs-lookup"><span data-stu-id="67946-548">Value</span></span>|
|---|---|
|[<span data-ttu-id="67946-549">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="67946-549">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="67946-550">预览</span><span class="sxs-lookup"><span data-stu-id="67946-550">Preview</span></span>|
|[<span data-ttu-id="67946-551">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="67946-551">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="67946-552">ReadItem</span><span class="sxs-lookup"><span data-stu-id="67946-552">ReadItem</span></span>|
|[<span data-ttu-id="67946-553">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="67946-553">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="67946-554">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="67946-554">Compose or read</span></span>|

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="67946-555">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="67946-555">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="67946-556">提供对事件的必需与会者的访问。</span><span class="sxs-lookup"><span data-stu-id="67946-556">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="67946-557">对象的类型和访问级别取决于当前项的模式。</span><span class="sxs-lookup"><span data-stu-id="67946-557">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="67946-558">阅读模式</span><span class="sxs-lookup"><span data-stu-id="67946-558">Read mode</span></span>

<span data-ttu-id="67946-559">`requiredAttendees` 属性返回一个数组，其中包含每个必需与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="67946-559">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="67946-560">撰写模式</span><span class="sxs-lookup"><span data-stu-id="67946-560">Compose mode</span></span>

<span data-ttu-id="67946-561">`requiredAttendees`属性返回`Recipients`对象，它提供用于获取或更新会议的必需的与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="67946-561">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="67946-562">类型：</span><span class="sxs-lookup"><span data-stu-id="67946-562">Type:</span></span>

*   <span data-ttu-id="67946-563">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="67946-563">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="67946-564">要求</span><span class="sxs-lookup"><span data-stu-id="67946-564">Requirements</span></span>

|<span data-ttu-id="67946-565">要求</span><span class="sxs-lookup"><span data-stu-id="67946-565">Requirement</span></span>|<span data-ttu-id="67946-566">值</span><span class="sxs-lookup"><span data-stu-id="67946-566">Value</span></span>|
|---|---|
|[<span data-ttu-id="67946-567">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="67946-567">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="67946-568">1.0</span><span class="sxs-lookup"><span data-stu-id="67946-568">1.0</span></span>|
|[<span data-ttu-id="67946-569">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="67946-569">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="67946-570">ReadItem</span><span class="sxs-lookup"><span data-stu-id="67946-570">ReadItem</span></span>|
|[<span data-ttu-id="67946-571">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="67946-571">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="67946-572">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="67946-572">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="67946-573">示例</span><span class="sxs-lookup"><span data-stu-id="67946-573">Example</span></span>

```
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails"></a><span data-ttu-id="67946-574">sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="67946-574">sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span></span>

<span data-ttu-id="67946-p126">获取电子邮件发件人的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="67946-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="67946-p127">[`from`](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) 和 `sender` 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="67946-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="67946-579">`recipientType`属性`EmailAddressDetails`对象在`sender`属性是`undefined`。</span><span class="sxs-lookup"><span data-stu-id="67946-579">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="67946-580">类型:</span><span class="sxs-lookup"><span data-stu-id="67946-580">Type:</span></span>

*   [<span data-ttu-id="67946-581">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="67946-581">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="67946-582">要求</span><span class="sxs-lookup"><span data-stu-id="67946-582">Requirements</span></span>

|<span data-ttu-id="67946-583">要求</span><span class="sxs-lookup"><span data-stu-id="67946-583">Requirement</span></span>|<span data-ttu-id="67946-584">值</span><span class="sxs-lookup"><span data-stu-id="67946-584">Value</span></span>|
|---|---|
|[<span data-ttu-id="67946-585">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="67946-585">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="67946-586">1.0</span><span class="sxs-lookup"><span data-stu-id="67946-586">1.0</span></span>|
|[<span data-ttu-id="67946-587">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="67946-587">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="67946-588">ReadItem</span><span class="sxs-lookup"><span data-stu-id="67946-588">ReadItem</span></span>|
|[<span data-ttu-id="67946-589">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="67946-589">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="67946-590">阅读</span><span class="sxs-lookup"><span data-stu-id="67946-590">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="67946-591">示例</span><span class="sxs-lookup"><span data-stu-id="67946-591">Example</span></span>

```
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="67946-592">(nullable) 可能指向： 字符串</span><span class="sxs-lookup"><span data-stu-id="67946-592">(nullable) seriesId :String</span></span>

<span data-ttu-id="67946-593">获取实例所属的系列的 id。</span><span class="sxs-lookup"><span data-stu-id="67946-593">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="67946-594">在 OWA 和 Outlook 中，`seriesId`返回父 （系列） 项此项目所属的 Exchange Web Services (EWS) ID。</span><span class="sxs-lookup"><span data-stu-id="67946-594">In OWA and Outlook, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="67946-595">但是，在 iOS 和 Android，`seriesId`返回父项的 REST ID。</span><span class="sxs-lookup"><span data-stu-id="67946-595">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="67946-596">`seriesId` 属性返回的标识符与 Exchange Web 服务项目标识符相同。</span><span class="sxs-lookup"><span data-stu-id="67946-596">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="67946-597">`seriesId`属性不是与 Outlook Id 使用 Outlook REST API 相同。</span><span class="sxs-lookup"><span data-stu-id="67946-597">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="67946-598">使用此值的 REST API 调用之前，应将其转换使用[Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string)。</span><span class="sxs-lookup"><span data-stu-id="67946-598">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="67946-599">有关详细信息，请参阅[使用 Outlook REST Api 从 Outlook 外接程序](https://docs.microsoft.com/outlook/add-ins/use-rest-api)。</span><span class="sxs-lookup"><span data-stu-id="67946-599">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="67946-600">`seriesId`属性返回`null`如不具有父项的项单个系列项目、 约会或会议请求，并返回`undefined`不会议请求的其他任何项。</span><span class="sxs-lookup"><span data-stu-id="67946-600">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="67946-601">类型:</span><span class="sxs-lookup"><span data-stu-id="67946-601">Type:</span></span>

* <span data-ttu-id="67946-602">String</span><span class="sxs-lookup"><span data-stu-id="67946-602">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="67946-603">要求</span><span class="sxs-lookup"><span data-stu-id="67946-603">Requirements</span></span>

|<span data-ttu-id="67946-604">要求</span><span class="sxs-lookup"><span data-stu-id="67946-604">Requirement</span></span>|<span data-ttu-id="67946-605">值</span><span class="sxs-lookup"><span data-stu-id="67946-605">Value</span></span>|
|---|---|
|[<span data-ttu-id="67946-606">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="67946-606">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="67946-607">预览</span><span class="sxs-lookup"><span data-stu-id="67946-607">Preview</span></span>|
|[<span data-ttu-id="67946-608">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="67946-608">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="67946-609">ReadItem</span><span class="sxs-lookup"><span data-stu-id="67946-609">ReadItem</span></span>|
|[<span data-ttu-id="67946-610">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="67946-610">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="67946-611">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="67946-611">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="67946-612">示例</span><span class="sxs-lookup"><span data-stu-id="67946-612">Example</span></span>

```
var seriesId = Office.context.mailbox.item.seriesId; 
var isSeries = (seriesId == null);
```

####  <a name="start-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="67946-613">start :Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="67946-613">start :Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="67946-614">获取或设置约会开始的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="67946-614">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="67946-p130">将 `start` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) 方法将该值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="67946-p130">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="67946-617">阅读模式</span><span class="sxs-lookup"><span data-stu-id="67946-617">Read mode</span></span>

<span data-ttu-id="67946-618">`start` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="67946-618">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="67946-619">撰写模式</span><span class="sxs-lookup"><span data-stu-id="67946-619">Compose mode</span></span>

<span data-ttu-id="67946-620">`start` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="67946-620">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="67946-621">使用 [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) 方法设置开始时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="67946-621">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="67946-622">类型：</span><span class="sxs-lookup"><span data-stu-id="67946-622">Type:</span></span>

*   <span data-ttu-id="67946-623">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="67946-623">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="67946-624">要求</span><span class="sxs-lookup"><span data-stu-id="67946-624">Requirements</span></span>

|<span data-ttu-id="67946-625">要求</span><span class="sxs-lookup"><span data-stu-id="67946-625">Requirement</span></span>|<span data-ttu-id="67946-626">值</span><span class="sxs-lookup"><span data-stu-id="67946-626">Value</span></span>|
|---|---|
|[<span data-ttu-id="67946-627">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="67946-627">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="67946-628">1.0</span><span class="sxs-lookup"><span data-stu-id="67946-628">1.0</span></span>|
|[<span data-ttu-id="67946-629">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="67946-629">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="67946-630">ReadItem</span><span class="sxs-lookup"><span data-stu-id="67946-630">ReadItem</span></span>|
|[<span data-ttu-id="67946-631">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="67946-631">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="67946-632">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="67946-632">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="67946-633">示例</span><span class="sxs-lookup"><span data-stu-id="67946-633">Example</span></span>

<span data-ttu-id="67946-634">以下示例通过使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) 方法，设置撰写模式下约会的开始时间。</span><span class="sxs-lookup"><span data-stu-id="67946-634">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

####  <a name="subject-stringsubjectjavascriptapioutlookofficesubject"></a><span data-ttu-id="67946-635">subject :String|[Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="67946-635">subject :String|[Subject](/javascript/api/outlook/office.subject)</span></span>

<span data-ttu-id="67946-636">获取或设置显示在项目的主题字段中的说明。</span><span class="sxs-lookup"><span data-stu-id="67946-636">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="67946-637">`subject` 属性获取或设置由电子邮件服务器发送项目时的整个主题。</span><span class="sxs-lookup"><span data-stu-id="67946-637">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="67946-638">阅读模式</span><span class="sxs-lookup"><span data-stu-id="67946-638">Read mode</span></span>

<span data-ttu-id="67946-p131">`subject` 属性返回一个字符串。使用 [`normalizedSubject`](#normalizedsubject-string) 属性获取不带任何前导前缀（如 `RE:` 和 `FW:`）的主题。</span><span class="sxs-lookup"><span data-stu-id="67946-p131">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="67946-641">撰写模式</span><span class="sxs-lookup"><span data-stu-id="67946-641">Compose mode</span></span>

<span data-ttu-id="67946-642">`subject` 属性返回一个 `Subject` 对象，该对象提供用于获取和设置主题的方法。</span><span class="sxs-lookup"><span data-stu-id="67946-642">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="67946-643">类型：</span><span class="sxs-lookup"><span data-stu-id="67946-643">Type:</span></span>

*   <span data-ttu-id="67946-644">String | [Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="67946-644">String | [Subject](/javascript/api/outlook/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="67946-645">要求</span><span class="sxs-lookup"><span data-stu-id="67946-645">Requirements</span></span>

|<span data-ttu-id="67946-646">要求</span><span class="sxs-lookup"><span data-stu-id="67946-646">Requirement</span></span>|<span data-ttu-id="67946-647">值</span><span class="sxs-lookup"><span data-stu-id="67946-647">Value</span></span>|
|---|---|
|[<span data-ttu-id="67946-648">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="67946-648">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="67946-649">1.0</span><span class="sxs-lookup"><span data-stu-id="67946-649">1.0</span></span>|
|[<span data-ttu-id="67946-650">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="67946-650">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="67946-651">ReadItem</span><span class="sxs-lookup"><span data-stu-id="67946-651">ReadItem</span></span>|
|[<span data-ttu-id="67946-652">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="67946-652">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="67946-653">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="67946-653">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="67946-654">到： 数组。 <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[收件人](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="67946-654">to :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="67946-655">提供对邮件的**到**行上的收件人访问。</span><span class="sxs-lookup"><span data-stu-id="67946-655">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="67946-656">对象的类型和访问级别取决于当前项的模式。</span><span class="sxs-lookup"><span data-stu-id="67946-656">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="67946-657">阅读模式</span><span class="sxs-lookup"><span data-stu-id="67946-657">Read mode</span></span>

<span data-ttu-id="67946-p133">`to` 属性返回包含邮件的**收件人**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="67946-p133">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="67946-660">撰写模式</span><span class="sxs-lookup"><span data-stu-id="67946-660">Compose mode</span></span>

<span data-ttu-id="67946-661">`to`属性返回`Recipients`提供方法用于获取或更新邮件的**到**行上的收件人对象。</span><span class="sxs-lookup"><span data-stu-id="67946-661">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="67946-662">类型：</span><span class="sxs-lookup"><span data-stu-id="67946-662">Type:</span></span>

*   <span data-ttu-id="67946-663">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="67946-663">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="67946-664">要求</span><span class="sxs-lookup"><span data-stu-id="67946-664">Requirements</span></span>

|<span data-ttu-id="67946-665">要求</span><span class="sxs-lookup"><span data-stu-id="67946-665">Requirement</span></span>|<span data-ttu-id="67946-666">值</span><span class="sxs-lookup"><span data-stu-id="67946-666">Value</span></span>|
|---|---|
|[<span data-ttu-id="67946-667">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="67946-667">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="67946-668">1.0</span><span class="sxs-lookup"><span data-stu-id="67946-668">1.0</span></span>|
|[<span data-ttu-id="67946-669">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="67946-669">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="67946-670">ReadItem</span><span class="sxs-lookup"><span data-stu-id="67946-670">ReadItem</span></span>|
|[<span data-ttu-id="67946-671">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="67946-671">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="67946-672">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="67946-672">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="67946-673">示例</span><span class="sxs-lookup"><span data-stu-id="67946-673">Example</span></span>

```
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="67946-674">方法</span><span class="sxs-lookup"><span data-stu-id="67946-674">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="67946-675">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="67946-675">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="67946-676">将文件作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="67946-676">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="67946-677">`addFileAttachmentAsync` 方法在指定的 URI 上载文件并将其附加到撰写窗体中的项目。</span><span class="sxs-lookup"><span data-stu-id="67946-677">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="67946-678">你随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="67946-678">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="67946-679">参数：</span><span class="sxs-lookup"><span data-stu-id="67946-679">Parameters:</span></span>
|<span data-ttu-id="67946-680">名称</span><span class="sxs-lookup"><span data-stu-id="67946-680">Name</span></span>|<span data-ttu-id="67946-681">类型</span><span class="sxs-lookup"><span data-stu-id="67946-681">Type</span></span>|<span data-ttu-id="67946-682">属性</span><span class="sxs-lookup"><span data-stu-id="67946-682">Attributes</span></span>|<span data-ttu-id="67946-683">说明</span><span class="sxs-lookup"><span data-stu-id="67946-683">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="67946-684">字符串</span><span class="sxs-lookup"><span data-stu-id="67946-684">String</span></span>||<span data-ttu-id="67946-p134">提供附加到邮件或约会的文件的位置的 URI。最大长度为 2048 个字符。</span><span class="sxs-lookup"><span data-stu-id="67946-p134">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="67946-687">字符串</span><span class="sxs-lookup"><span data-stu-id="67946-687">String</span></span>||<span data-ttu-id="67946-p135">在附件上载过程中显示的附件名称。最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="67946-p135">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="67946-690">对象</span><span class="sxs-lookup"><span data-stu-id="67946-690">Object</span></span>|<span data-ttu-id="67946-691">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="67946-691">&lt;optional&gt;</span></span>|<span data-ttu-id="67946-692">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="67946-692">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="67946-693">对象</span><span class="sxs-lookup"><span data-stu-id="67946-693">Object</span></span>|<span data-ttu-id="67946-694">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="67946-694">&lt;optional&gt;</span></span>|<span data-ttu-id="67946-695">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="67946-695">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="67946-696">布尔值</span><span class="sxs-lookup"><span data-stu-id="67946-696">Boolean</span></span>|<span data-ttu-id="67946-697">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="67946-697">&lt;optional&gt;</span></span>|<span data-ttu-id="67946-698">如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="67946-698">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="67946-699">函数</span><span class="sxs-lookup"><span data-stu-id="67946-699">function</span></span>|<span data-ttu-id="67946-700">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="67946-700">&lt;optional&gt;</span></span>|<span data-ttu-id="67946-701">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="67946-701">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="67946-702">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="67946-702">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="67946-703">如果上传附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="67946-703">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="67946-704">错误</span><span class="sxs-lookup"><span data-stu-id="67946-704">Errors</span></span>

|<span data-ttu-id="67946-705">错误代码</span><span class="sxs-lookup"><span data-stu-id="67946-705">Error code</span></span>|<span data-ttu-id="67946-706">说明</span><span class="sxs-lookup"><span data-stu-id="67946-706">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="67946-707">附件大小超过了允许的大小。</span><span class="sxs-lookup"><span data-stu-id="67946-707">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="67946-708">该附件的扩展名不是允许的扩展名。</span><span class="sxs-lookup"><span data-stu-id="67946-708">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="67946-709">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="67946-709">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="67946-710">要求</span><span class="sxs-lookup"><span data-stu-id="67946-710">Requirements</span></span>

|<span data-ttu-id="67946-711">要求</span><span class="sxs-lookup"><span data-stu-id="67946-711">Requirement</span></span>|<span data-ttu-id="67946-712">值</span><span class="sxs-lookup"><span data-stu-id="67946-712">Value</span></span>|
|---|---|
|[<span data-ttu-id="67946-713">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="67946-713">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="67946-714">1.1</span><span class="sxs-lookup"><span data-stu-id="67946-714">1.1</span></span>|
|[<span data-ttu-id="67946-715">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="67946-715">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="67946-716">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="67946-716">ReadWriteItem</span></span>|
|[<span data-ttu-id="67946-717">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="67946-717">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="67946-718">撰写</span><span class="sxs-lookup"><span data-stu-id="67946-718">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="67946-719">示例</span><span class="sxs-lookup"><span data-stu-id="67946-719">Examples</span></span>

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

<span data-ttu-id="67946-720">以下示例将图像文件添加为内联附件，并在邮件正文中引用该附件。</span><span class="sxs-lookup"><span data-stu-id="67946-720">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a><span data-ttu-id="67946-721">addFileAttachmentFromBase64Async (base64File，attachmentName，[选项] [回调])</span><span class="sxs-lookup"><span data-stu-id="67946-721">addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="67946-722">添加从 base64 编码到邮件或约会作为附件的文件。</span><span class="sxs-lookup"><span data-stu-id="67946-722">Adds a file from the base64 encoding to a message or appointment as an attachment.</span></span>

<span data-ttu-id="67946-723">`addFileAttachmentFromBase64Async`方法从 base64 编码将文件上载，并将其附加到撰写窗体中的项。</span><span class="sxs-lookup"><span data-stu-id="67946-723">The `addFileAttachmentFromBase64Async` method uploads the file from the base64 encoding and attaches it to the item in the compose form.</span></span> <span data-ttu-id="67946-724">此方法返回 AsyncResult.value 对象中的附件标识符。</span><span class="sxs-lookup"><span data-stu-id="67946-724">This method returns the attachment identifier in the AsyncResult.value object.</span></span>

<span data-ttu-id="67946-725">你随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="67946-725">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="67946-726">参数：</span><span class="sxs-lookup"><span data-stu-id="67946-726">Parameters:</span></span>
|<span data-ttu-id="67946-727">名称</span><span class="sxs-lookup"><span data-stu-id="67946-727">Name</span></span>|<span data-ttu-id="67946-728">类型</span><span class="sxs-lookup"><span data-stu-id="67946-728">Type</span></span>|<span data-ttu-id="67946-729">属性</span><span class="sxs-lookup"><span data-stu-id="67946-729">Attributes</span></span>|<span data-ttu-id="67946-730">说明</span><span class="sxs-lookup"><span data-stu-id="67946-730">Description</span></span>|
|---|---|---|---|
|`base64File`|<span data-ttu-id="67946-731">String</span><span class="sxs-lookup"><span data-stu-id="67946-731">String</span></span>||<span data-ttu-id="67946-732">以 base64 编码的图像或文件添加到电子邮件或事件的内容。</span><span class="sxs-lookup"><span data-stu-id="67946-732">The base64 encoded content of an image or file to be added to an email or event.</span></span>|
|`attachmentName`|<span data-ttu-id="67946-733">字符串</span><span class="sxs-lookup"><span data-stu-id="67946-733">String</span></span>||<span data-ttu-id="67946-p137">在附件上载过程中显示的附件名称。最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="67946-p137">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="67946-736">对象</span><span class="sxs-lookup"><span data-stu-id="67946-736">Object</span></span>|<span data-ttu-id="67946-737">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="67946-737">&lt;optional&gt;</span></span>|<span data-ttu-id="67946-738">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="67946-738">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="67946-739">对象</span><span class="sxs-lookup"><span data-stu-id="67946-739">Object</span></span>|<span data-ttu-id="67946-740">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="67946-740">&lt;optional&gt;</span></span>|<span data-ttu-id="67946-741">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="67946-741">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="67946-742">布尔值</span><span class="sxs-lookup"><span data-stu-id="67946-742">Boolean</span></span>|<span data-ttu-id="67946-743">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="67946-743">&lt;optional&gt;</span></span>|<span data-ttu-id="67946-744">如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="67946-744">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="67946-745">函数</span><span class="sxs-lookup"><span data-stu-id="67946-745">function</span></span>|<span data-ttu-id="67946-746">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="67946-746">&lt;optional&gt;</span></span>|<span data-ttu-id="67946-747">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="67946-747">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="67946-748">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="67946-748">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="67946-749">如果上传附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="67946-749">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="67946-750">错误</span><span class="sxs-lookup"><span data-stu-id="67946-750">Errors</span></span>

|<span data-ttu-id="67946-751">错误代码</span><span class="sxs-lookup"><span data-stu-id="67946-751">Error code</span></span>|<span data-ttu-id="67946-752">说明</span><span class="sxs-lookup"><span data-stu-id="67946-752">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="67946-753">附件大小超过了允许的大小。</span><span class="sxs-lookup"><span data-stu-id="67946-753">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="67946-754">该附件的扩展名不是允许的扩展名。</span><span class="sxs-lookup"><span data-stu-id="67946-754">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="67946-755">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="67946-755">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="67946-756">要求</span><span class="sxs-lookup"><span data-stu-id="67946-756">Requirements</span></span>

|<span data-ttu-id="67946-757">要求</span><span class="sxs-lookup"><span data-stu-id="67946-757">Requirement</span></span>|<span data-ttu-id="67946-758">值</span><span class="sxs-lookup"><span data-stu-id="67946-758">Value</span></span>|
|---|---|
|[<span data-ttu-id="67946-759">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="67946-759">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="67946-760">预览</span><span class="sxs-lookup"><span data-stu-id="67946-760">Preview</span></span>|
|[<span data-ttu-id="67946-761">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="67946-761">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="67946-762">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="67946-762">ReadWriteItem</span></span>|
|[<span data-ttu-id="67946-763">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="67946-763">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="67946-764">撰写</span><span class="sxs-lookup"><span data-stu-id="67946-764">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="67946-765">示例</span><span class="sxs-lookup"><span data-stu-id="67946-765">Examples</span></span>

```js
Office.context.mailbox.item.addFileAttachmentFromBase64Async(
  base64String,
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

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="67946-766">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="67946-766">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="67946-767">添加支持事件的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="67946-767">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="67946-768">当前受支持的事件类型的`Office.EventType.AppointmentTimeChanged`， `Office.EventType.RecipientsChanged`，和`Office.EventType.RecurrencePatternChanged`</span><span class="sxs-lookup"><span data-stu-id="67946-768">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrencePatternChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="67946-769">参数：</span><span class="sxs-lookup"><span data-stu-id="67946-769">Parameters:</span></span>

| <span data-ttu-id="67946-770">名称</span><span class="sxs-lookup"><span data-stu-id="67946-770">Name</span></span> | <span data-ttu-id="67946-771">类型</span><span class="sxs-lookup"><span data-stu-id="67946-771">Type</span></span> | <span data-ttu-id="67946-772">属性</span><span class="sxs-lookup"><span data-stu-id="67946-772">Attributes</span></span> | <span data-ttu-id="67946-773">说明</span><span class="sxs-lookup"><span data-stu-id="67946-773">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="67946-774">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="67946-774">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="67946-775">应调用处理程序的事件。</span><span class="sxs-lookup"><span data-stu-id="67946-775">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="67946-776">函数</span><span class="sxs-lookup"><span data-stu-id="67946-776">Function</span></span> || <span data-ttu-id="67946-p138">用于处理事件的函数。此函数必须接受一个参数，即对象文本。参数上的 `type` 属性将匹配传递给 `addHandlerAsync` 的 `eventType` 参数。</span><span class="sxs-lookup"><span data-stu-id="67946-p138">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="67946-780">Object</span><span class="sxs-lookup"><span data-stu-id="67946-780">Object</span></span> | <span data-ttu-id="67946-781">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="67946-781">&lt;optional&gt;</span></span> | <span data-ttu-id="67946-782">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="67946-782">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="67946-783">对象</span><span class="sxs-lookup"><span data-stu-id="67946-783">Object</span></span> | <span data-ttu-id="67946-784">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="67946-784">&lt;optional&gt;</span></span> | <span data-ttu-id="67946-785">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="67946-785">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="67946-786">函数</span><span class="sxs-lookup"><span data-stu-id="67946-786">function</span></span>| <span data-ttu-id="67946-787">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="67946-787">&lt;optional&gt;</span></span>|<span data-ttu-id="67946-788">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="67946-788">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="67946-789">Requirements</span><span class="sxs-lookup"><span data-stu-id="67946-789">Requirements</span></span>

|<span data-ttu-id="67946-790">要求</span><span class="sxs-lookup"><span data-stu-id="67946-790">Requirement</span></span>| <span data-ttu-id="67946-791">值</span><span class="sxs-lookup"><span data-stu-id="67946-791">Value</span></span>|
|---|---|
|[<span data-ttu-id="67946-792">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="67946-792">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="67946-793">预览</span><span class="sxs-lookup"><span data-stu-id="67946-793">Preview</span></span> |
|[<span data-ttu-id="67946-794">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="67946-794">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="67946-795">ReadItem</span><span class="sxs-lookup"><span data-stu-id="67946-795">ReadItem</span></span> |
|[<span data-ttu-id="67946-796">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="67946-796">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="67946-797">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="67946-797">Compose or read</span></span> |

##### <a name="example"></a><span data-ttu-id="67946-798">示例</span><span class="sxs-lookup"><span data-stu-id="67946-798">Example</span></span>

```
Office.initialize = function (reason) {
  $(document).ready(function () {
    Office.context.mailbox.item.addHandlerAsync(Office.EventType.RecurrencePatternChanged, loadNewItem, function (result) {
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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="67946-799">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="67946-799">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="67946-800">将 Exchange 项目（如邮件）作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="67946-800">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="67946-p139">`addItemAttachmentAsync` 方法将包含指定 Exchange 标识符的项目附加到撰写窗体中的项目。如果指定一个回调方法，此方法使用 `asyncResult` 参数调用，该参数包含一个附件标识符或代码，指示附加项目过程中出现的任何错误。可以使用 `options` 参数将状态信息传递给回调方法（如果需要）。</span><span class="sxs-lookup"><span data-stu-id="67946-p139">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="67946-804">你随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="67946-804">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="67946-805">如果您 Office 加载项运行的 Outlook Web App 中`addItemAttachmentAsync`方法可以附加到正在编辑; 的项目之外的项目的项目但是，这不受支持，并且不建议使用。</span><span class="sxs-lookup"><span data-stu-id="67946-805">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="67946-806">参数：</span><span class="sxs-lookup"><span data-stu-id="67946-806">Parameters:</span></span>

|<span data-ttu-id="67946-807">名称</span><span class="sxs-lookup"><span data-stu-id="67946-807">Name</span></span>|<span data-ttu-id="67946-808">类型</span><span class="sxs-lookup"><span data-stu-id="67946-808">Type</span></span>|<span data-ttu-id="67946-809">属性</span><span class="sxs-lookup"><span data-stu-id="67946-809">Attributes</span></span>|<span data-ttu-id="67946-810">说明</span><span class="sxs-lookup"><span data-stu-id="67946-810">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="67946-811">字符串</span><span class="sxs-lookup"><span data-stu-id="67946-811">String</span></span>||<span data-ttu-id="67946-p140">要附加的项目的 Exchange 标识符。最大长度为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="67946-p140">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="67946-814">字符串</span><span class="sxs-lookup"><span data-stu-id="67946-814">String</span></span>||<span data-ttu-id="67946-p141">要附加的项目的主题。最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="67946-p141">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="67946-817">对象</span><span class="sxs-lookup"><span data-stu-id="67946-817">Object</span></span>|<span data-ttu-id="67946-818">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="67946-818">&lt;optional&gt;</span></span>|<span data-ttu-id="67946-819">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="67946-819">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="67946-820">对象</span><span class="sxs-lookup"><span data-stu-id="67946-820">Object</span></span>|<span data-ttu-id="67946-821">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="67946-821">&lt;optional&gt;</span></span>|<span data-ttu-id="67946-822">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="67946-822">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="67946-823">函数</span><span class="sxs-lookup"><span data-stu-id="67946-823">function</span></span>|<span data-ttu-id="67946-824">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="67946-824">&lt;optional&gt;</span></span>|<span data-ttu-id="67946-825">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="67946-825">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="67946-826">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="67946-826">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="67946-827">如果添加附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="67946-827">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="67946-828">错误</span><span class="sxs-lookup"><span data-stu-id="67946-828">Errors</span></span>

|<span data-ttu-id="67946-829">错误代码</span><span class="sxs-lookup"><span data-stu-id="67946-829">Error code</span></span>|<span data-ttu-id="67946-830">说明</span><span class="sxs-lookup"><span data-stu-id="67946-830">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="67946-831">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="67946-831">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="67946-832">要求</span><span class="sxs-lookup"><span data-stu-id="67946-832">Requirements</span></span>

|<span data-ttu-id="67946-833">要求</span><span class="sxs-lookup"><span data-stu-id="67946-833">Requirement</span></span>|<span data-ttu-id="67946-834">值</span><span class="sxs-lookup"><span data-stu-id="67946-834">Value</span></span>|
|---|---|
|[<span data-ttu-id="67946-835">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="67946-835">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="67946-836">1.1</span><span class="sxs-lookup"><span data-stu-id="67946-836">1.1</span></span>|
|[<span data-ttu-id="67946-837">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="67946-837">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="67946-838">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="67946-838">ReadWriteItem</span></span>|
|[<span data-ttu-id="67946-839">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="67946-839">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="67946-840">撰写</span><span class="sxs-lookup"><span data-stu-id="67946-840">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="67946-841">示例</span><span class="sxs-lookup"><span data-stu-id="67946-841">Example</span></span>

<span data-ttu-id="67946-842">以下示例将现有的 Outlook 项目添加为名为 `My Attachment` 的附件。</span><span class="sxs-lookup"><span data-stu-id="67946-842">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

####  <a name="close"></a><span data-ttu-id="67946-843">close()</span><span class="sxs-lookup"><span data-stu-id="67946-843">close()</span></span>

<span data-ttu-id="67946-844">关闭当前正在撰写的项目。</span><span class="sxs-lookup"><span data-stu-id="67946-844">Closes the current item that is being composed.</span></span>

<span data-ttu-id="67946-p142">`close` 方法的行为取决于要撰写的项目的当前状态。如果项目具有未保存的更改，客户端将提示用户保存、放弃或取消关闭操作。</span><span class="sxs-lookup"><span data-stu-id="67946-p142">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="67946-847">在 Outlook 中的 web，如果该项目是约会和它之前已保存使用在`saveAsync`、 提示用户保存、 放弃，或取消，即使未发生更改由于项目上次保存。</span><span class="sxs-lookup"><span data-stu-id="67946-847">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="67946-848">在 Outlook 桌面客户端中，如果邮件是内联答复，`close` 方法不起作用。</span><span class="sxs-lookup"><span data-stu-id="67946-848">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="67946-849">要求</span><span class="sxs-lookup"><span data-stu-id="67946-849">Requirements</span></span>

|<span data-ttu-id="67946-850">要求</span><span class="sxs-lookup"><span data-stu-id="67946-850">Requirement</span></span>|<span data-ttu-id="67946-851">值</span><span class="sxs-lookup"><span data-stu-id="67946-851">Value</span></span>|
|---|---|
|[<span data-ttu-id="67946-852">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="67946-852">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="67946-853">1.3</span><span class="sxs-lookup"><span data-stu-id="67946-853">1.3</span></span>|
|[<span data-ttu-id="67946-854">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="67946-854">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="67946-855">受限</span><span class="sxs-lookup"><span data-stu-id="67946-855">Restricted</span></span>|
|[<span data-ttu-id="67946-856">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="67946-856">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="67946-857">撰写</span><span class="sxs-lookup"><span data-stu-id="67946-857">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="67946-858">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="67946-858">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="67946-859">显示答复窗体，其中包括所选邮件的发件人和所有收件人或所选约会的组织者和所有与会者。</span><span class="sxs-lookup"><span data-stu-id="67946-859">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="67946-860">在适用于 iOS 的 Outlook 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="67946-860">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="67946-861">在 Outlook Web App 中，答复窗体显示为包含 3 列视图的弹出式窗体以及包含 2 列或 1 列视图的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="67946-861">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="67946-862">如果任意字符串参数超出其限制，`displayReplyAllForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="67946-862">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="67946-p143">当在 `formData.attachments` 参数中指定附件时，Outlook 和 Outlook Web App 尝试下载所有附件并将其附加到答复窗体。如果无法添加任何附件，则在窗体 UI 中显示错误。如果这不可能，则不引发错误消息。</span><span class="sxs-lookup"><span data-stu-id="67946-p143">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="67946-866">参数：</span><span class="sxs-lookup"><span data-stu-id="67946-866">Parameters:</span></span>

|<span data-ttu-id="67946-867">名称</span><span class="sxs-lookup"><span data-stu-id="67946-867">Name</span></span>|<span data-ttu-id="67946-868">类型</span><span class="sxs-lookup"><span data-stu-id="67946-868">Type</span></span>|<span data-ttu-id="67946-869">属性</span><span class="sxs-lookup"><span data-stu-id="67946-869">Attributes</span></span>|<span data-ttu-id="67946-870">说明</span><span class="sxs-lookup"><span data-stu-id="67946-870">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="67946-871">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="67946-871">String &#124; Object</span></span>||<span data-ttu-id="67946-p144">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="67946-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="67946-874">**或**</span><span class="sxs-lookup"><span data-stu-id="67946-874">**OR**</span></span><br/><span data-ttu-id="67946-p145">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="67946-p145">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="67946-877">字符串</span><span class="sxs-lookup"><span data-stu-id="67946-877">String</span></span>|<span data-ttu-id="67946-878">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="67946-878">&lt;optional&gt;</span></span>|<span data-ttu-id="67946-p146">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="67946-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="67946-881">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="67946-881">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="67946-882">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="67946-882">&lt;optional&gt;</span></span>|<span data-ttu-id="67946-883">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="67946-883">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="67946-884">字符串</span><span class="sxs-lookup"><span data-stu-id="67946-884">String</span></span>||<span data-ttu-id="67946-p147">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="67946-p147">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="67946-887">字符串</span><span class="sxs-lookup"><span data-stu-id="67946-887">String</span></span>||<span data-ttu-id="67946-888">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="67946-888">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="67946-889">String</span><span class="sxs-lookup"><span data-stu-id="67946-889">String</span></span>||<span data-ttu-id="67946-p148">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="67946-p148">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="67946-892">Boolean</span><span class="sxs-lookup"><span data-stu-id="67946-892">Boolean</span></span>||<span data-ttu-id="67946-p149">仅在将 `type` 设置为 `file` 时使用。如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="67946-p149">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="67946-895">String</span><span class="sxs-lookup"><span data-stu-id="67946-895">String</span></span>||<span data-ttu-id="67946-p150">仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="67946-p150">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="67946-899">函数</span><span class="sxs-lookup"><span data-stu-id="67946-899">function</span></span>|<span data-ttu-id="67946-900">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="67946-900">&lt;optional&gt;</span></span>|<span data-ttu-id="67946-901">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="67946-901">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="67946-902">要求</span><span class="sxs-lookup"><span data-stu-id="67946-902">Requirements</span></span>

|<span data-ttu-id="67946-903">要求</span><span class="sxs-lookup"><span data-stu-id="67946-903">Requirement</span></span>|<span data-ttu-id="67946-904">值</span><span class="sxs-lookup"><span data-stu-id="67946-904">Value</span></span>|
|---|---|
|[<span data-ttu-id="67946-905">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="67946-905">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="67946-906">1.0</span><span class="sxs-lookup"><span data-stu-id="67946-906">1.0</span></span>|
|[<span data-ttu-id="67946-907">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="67946-907">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="67946-908">ReadItem</span><span class="sxs-lookup"><span data-stu-id="67946-908">ReadItem</span></span>|
|[<span data-ttu-id="67946-909">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="67946-909">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="67946-910">阅读</span><span class="sxs-lookup"><span data-stu-id="67946-910">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="67946-911">示例</span><span class="sxs-lookup"><span data-stu-id="67946-911">Examples</span></span>

<span data-ttu-id="67946-912">以下代码将一个字符串传递到 `displayReplyAllForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="67946-912">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="67946-913">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="67946-913">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="67946-914">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="67946-914">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="67946-915">使用正文和文件附件答复。</span><span class="sxs-lookup"><span data-stu-id="67946-915">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="67946-916">使用正文和项目附件答复。</span><span class="sxs-lookup"><span data-stu-id="67946-916">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="67946-917">使用正文、文件附件、项目附件和回调答复。</span><span class="sxs-lookup"><span data-stu-id="67946-917">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="67946-918">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="67946-918">displayReplyForm(formData)</span></span>

<span data-ttu-id="67946-919">显示答复窗体，其中仅包括所选邮件的发件人或所选约会的组织者。</span><span class="sxs-lookup"><span data-stu-id="67946-919">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="67946-920">在适用于 iOS 的 Outlook 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="67946-920">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="67946-921">在 Outlook Web App 中，答复窗体显示为包含 3 列视图的弹出式窗体以及包含 2 列或 1 列视图的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="67946-921">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="67946-922">如果任意字符串参数超出其限制，`displayReplyForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="67946-922">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="67946-p151">当在 `formData.attachments` 参数中指定附件时，Outlook 和 Outlook Web App 尝试下载所有附件并将其附加到答复窗体。如果无法添加任何附件，则在窗体 UI 中显示错误。如果这不可能，则不引发错误消息。</span><span class="sxs-lookup"><span data-stu-id="67946-p151">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="67946-926">参数：</span><span class="sxs-lookup"><span data-stu-id="67946-926">Parameters:</span></span>

|<span data-ttu-id="67946-927">名称</span><span class="sxs-lookup"><span data-stu-id="67946-927">Name</span></span>|<span data-ttu-id="67946-928">类型</span><span class="sxs-lookup"><span data-stu-id="67946-928">Type</span></span>|<span data-ttu-id="67946-929">属性</span><span class="sxs-lookup"><span data-stu-id="67946-929">Attributes</span></span>|<span data-ttu-id="67946-930">说明</span><span class="sxs-lookup"><span data-stu-id="67946-930">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="67946-931">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="67946-931">String &#124; Object</span></span>||<span data-ttu-id="67946-p152">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="67946-p152">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="67946-934">**或**</span><span class="sxs-lookup"><span data-stu-id="67946-934">**OR**</span></span><br/><span data-ttu-id="67946-p153">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="67946-p153">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="67946-937">字符串</span><span class="sxs-lookup"><span data-stu-id="67946-937">String</span></span>|<span data-ttu-id="67946-938">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="67946-938">&lt;optional&gt;</span></span>|<span data-ttu-id="67946-p154">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="67946-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="67946-941">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="67946-941">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="67946-942">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="67946-942">&lt;optional&gt;</span></span>|<span data-ttu-id="67946-943">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="67946-943">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="67946-944">字符串</span><span class="sxs-lookup"><span data-stu-id="67946-944">String</span></span>||<span data-ttu-id="67946-p155">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="67946-p155">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="67946-947">字符串</span><span class="sxs-lookup"><span data-stu-id="67946-947">String</span></span>||<span data-ttu-id="67946-948">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="67946-948">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="67946-949">String</span><span class="sxs-lookup"><span data-stu-id="67946-949">String</span></span>||<span data-ttu-id="67946-p156">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="67946-p156">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="67946-952">Boolean</span><span class="sxs-lookup"><span data-stu-id="67946-952">Boolean</span></span>||<span data-ttu-id="67946-p157">仅在将 `type` 设置为 `file` 时使用。如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="67946-p157">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="67946-955">String</span><span class="sxs-lookup"><span data-stu-id="67946-955">String</span></span>||<span data-ttu-id="67946-p158">仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="67946-p158">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="67946-959">函数</span><span class="sxs-lookup"><span data-stu-id="67946-959">function</span></span>|<span data-ttu-id="67946-960">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="67946-960">&lt;optional&gt;</span></span>|<span data-ttu-id="67946-961">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="67946-961">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="67946-962">要求</span><span class="sxs-lookup"><span data-stu-id="67946-962">Requirements</span></span>

|<span data-ttu-id="67946-963">要求</span><span class="sxs-lookup"><span data-stu-id="67946-963">Requirement</span></span>|<span data-ttu-id="67946-964">值</span><span class="sxs-lookup"><span data-stu-id="67946-964">Value</span></span>|
|---|---|
|[<span data-ttu-id="67946-965">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="67946-965">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="67946-966">1.0</span><span class="sxs-lookup"><span data-stu-id="67946-966">1.0</span></span>|
|[<span data-ttu-id="67946-967">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="67946-967">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="67946-968">ReadItem</span><span class="sxs-lookup"><span data-stu-id="67946-968">ReadItem</span></span>|
|[<span data-ttu-id="67946-969">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="67946-969">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="67946-970">阅读</span><span class="sxs-lookup"><span data-stu-id="67946-970">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="67946-971">示例</span><span class="sxs-lookup"><span data-stu-id="67946-971">Examples</span></span>

<span data-ttu-id="67946-972">以下代码将一个字符串传递到 `displayReplyForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="67946-972">The following code passes a string to the `displayReplyForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="67946-973">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="67946-973">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="67946-974">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="67946-974">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="67946-975">使用正文和文件附件答复。</span><span class="sxs-lookup"><span data-stu-id="67946-975">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="67946-976">使用正文和项目附件答复。</span><span class="sxs-lookup"><span data-stu-id="67946-976">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="67946-977">使用正文、文件附件、项目附件和回调答复。</span><span class="sxs-lookup"><span data-stu-id="67946-977">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="67946-978">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="67946-978">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="67946-979">获取在选定的项的正文中找到的实体。</span><span class="sxs-lookup"><span data-stu-id="67946-979">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="67946-980">在适用于 iOS 的 Outlook 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="67946-980">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="67946-981">要求</span><span class="sxs-lookup"><span data-stu-id="67946-981">Requirements</span></span>

|<span data-ttu-id="67946-982">要求</span><span class="sxs-lookup"><span data-stu-id="67946-982">Requirement</span></span>|<span data-ttu-id="67946-983">值</span><span class="sxs-lookup"><span data-stu-id="67946-983">Value</span></span>|
|---|---|
|[<span data-ttu-id="67946-984">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="67946-984">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="67946-985">1.0</span><span class="sxs-lookup"><span data-stu-id="67946-985">1.0</span></span>|
|[<span data-ttu-id="67946-986">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="67946-986">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="67946-987">ReadItem</span><span class="sxs-lookup"><span data-stu-id="67946-987">ReadItem</span></span>|
|[<span data-ttu-id="67946-988">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="67946-988">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="67946-989">阅读</span><span class="sxs-lookup"><span data-stu-id="67946-989">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="67946-990">返回：</span><span class="sxs-lookup"><span data-stu-id="67946-990">Returns:</span></span>

<span data-ttu-id="67946-991">类型：[Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="67946-991">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="67946-992">示例</span><span class="sxs-lookup"><span data-stu-id="67946-992">Example</span></span>

<span data-ttu-id="67946-993">以下示例访问当前项的主体中的联系人实体。</span><span class="sxs-lookup"><span data-stu-id="67946-993">The following example accesses the contacts entities in the current item's body.</span></span>

```
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="67946-994">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="67946-994">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="67946-995">获取选定的项的正文中找到的指定的实体类型的所有实体的数组。</span><span class="sxs-lookup"><span data-stu-id="67946-995">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="67946-996">在适用于 iOS 的 Outlook 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="67946-996">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="67946-997">参数：</span><span class="sxs-lookup"><span data-stu-id="67946-997">Parameters:</span></span>

|<span data-ttu-id="67946-998">名称</span><span class="sxs-lookup"><span data-stu-id="67946-998">Name</span></span>|<span data-ttu-id="67946-999">类型</span><span class="sxs-lookup"><span data-stu-id="67946-999">Type</span></span>|<span data-ttu-id="67946-1000">说明</span><span class="sxs-lookup"><span data-stu-id="67946-1000">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="67946-1001">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="67946-1001">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype)|<span data-ttu-id="67946-1002">EntityType 枚举值之一。</span><span class="sxs-lookup"><span data-stu-id="67946-1002">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="67946-1003">要求</span><span class="sxs-lookup"><span data-stu-id="67946-1003">Requirements</span></span>

|<span data-ttu-id="67946-1004">要求</span><span class="sxs-lookup"><span data-stu-id="67946-1004">Requirement</span></span>|<span data-ttu-id="67946-1005">值</span><span class="sxs-lookup"><span data-stu-id="67946-1005">Value</span></span>|
|---|---|
|[<span data-ttu-id="67946-1006">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="67946-1006">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="67946-1007">1.0</span><span class="sxs-lookup"><span data-stu-id="67946-1007">1.0</span></span>|
|[<span data-ttu-id="67946-1008">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="67946-1008">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="67946-1009">受限</span><span class="sxs-lookup"><span data-stu-id="67946-1009">Restricted</span></span>|
|[<span data-ttu-id="67946-1010">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="67946-1010">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="67946-1011">阅读</span><span class="sxs-lookup"><span data-stu-id="67946-1011">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="67946-1012">返回：</span><span class="sxs-lookup"><span data-stu-id="67946-1012">Returns:</span></span>

<span data-ttu-id="67946-1013">如果在 `entityType` 中传递的值不是 `EntityType` 枚举的有效成员，该方法返回 null。</span><span class="sxs-lookup"><span data-stu-id="67946-1013">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="67946-1014">如果没有指定类型的实体的存在于项目的正文中，该方法将返回空数组。</span><span class="sxs-lookup"><span data-stu-id="67946-1014">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="67946-1015">否则，返回的数组中对象的类型取决于 `entityType` 参数中请求实体的类型。</span><span class="sxs-lookup"><span data-stu-id="67946-1015">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="67946-1016">当使用此方法的最低权限级别**受限**时，某些实体类型需要 **ReadItem** 才能进行访问，如下表中所指定。</span><span class="sxs-lookup"><span data-stu-id="67946-1016">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="67946-1017">`entityType` 的值</span><span class="sxs-lookup"><span data-stu-id="67946-1017">Value of `entityType`</span></span>|<span data-ttu-id="67946-1018">返回的数组中对象的类型</span><span class="sxs-lookup"><span data-stu-id="67946-1018">Type of objects in returned array</span></span>|<span data-ttu-id="67946-1019">所需权限级别</span><span class="sxs-lookup"><span data-stu-id="67946-1019">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="67946-1020">String</span><span class="sxs-lookup"><span data-stu-id="67946-1020">String</span></span>|<span data-ttu-id="67946-1021">**受限**</span><span class="sxs-lookup"><span data-stu-id="67946-1021">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="67946-1022">Contact</span><span class="sxs-lookup"><span data-stu-id="67946-1022">Contact</span></span>|<span data-ttu-id="67946-1023">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="67946-1023">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="67946-1024">String</span><span class="sxs-lookup"><span data-stu-id="67946-1024">String</span></span>|<span data-ttu-id="67946-1025">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="67946-1025">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="67946-1026">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="67946-1026">MeetingSuggestion</span></span>|<span data-ttu-id="67946-1027">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="67946-1027">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="67946-1028">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="67946-1028">PhoneNumber</span></span>|<span data-ttu-id="67946-1029">**受限**</span><span class="sxs-lookup"><span data-stu-id="67946-1029">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="67946-1030">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="67946-1030">TaskSuggestion</span></span>|<span data-ttu-id="67946-1031">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="67946-1031">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="67946-1032">String</span><span class="sxs-lookup"><span data-stu-id="67946-1032">String</span></span>|<span data-ttu-id="67946-1033">**受限**</span><span class="sxs-lookup"><span data-stu-id="67946-1033">**Restricted**</span></span>|

<span data-ttu-id="67946-1034">类型：Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="67946-1034">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="67946-1035">示例</span><span class="sxs-lookup"><span data-stu-id="67946-1035">Example</span></span>

<span data-ttu-id="67946-1036">下面的示例演示如何访问表示当前项目的正文中的邮政地址的字符串数组。</span><span class="sxs-lookup"><span data-stu-id="67946-1036">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="67946-1037">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="67946-1037">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="67946-1038">返回传递清单 XML 文件中定义的命名筛选器的所选项目中的已知实体。</span><span class="sxs-lookup"><span data-stu-id="67946-1038">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="67946-1039">在适用于 iOS 的 Outlook 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="67946-1039">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="67946-1040">`getFilteredEntitiesByName` 方法返回匹配在具有指定 `FilterName` 元素值的清单 XML 文件中的 [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) 规则元素中定义的正则表达式的实体。</span><span class="sxs-lookup"><span data-stu-id="67946-1040">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="67946-1041">参数：</span><span class="sxs-lookup"><span data-stu-id="67946-1041">Parameters:</span></span>

|<span data-ttu-id="67946-1042">名称</span><span class="sxs-lookup"><span data-stu-id="67946-1042">Name</span></span>|<span data-ttu-id="67946-1043">类型</span><span class="sxs-lookup"><span data-stu-id="67946-1043">Type</span></span>|<span data-ttu-id="67946-1044">说明</span><span class="sxs-lookup"><span data-stu-id="67946-1044">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="67946-1045">字符串</span><span class="sxs-lookup"><span data-stu-id="67946-1045">String</span></span>|<span data-ttu-id="67946-1046">定义筛选器匹配的 `ItemHasKnownEntity` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="67946-1046">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="67946-1047">要求</span><span class="sxs-lookup"><span data-stu-id="67946-1047">Requirements</span></span>

|<span data-ttu-id="67946-1048">要求</span><span class="sxs-lookup"><span data-stu-id="67946-1048">Requirement</span></span>|<span data-ttu-id="67946-1049">值</span><span class="sxs-lookup"><span data-stu-id="67946-1049">Value</span></span>|
|---|---|
|[<span data-ttu-id="67946-1050">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="67946-1050">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="67946-1051">1.0</span><span class="sxs-lookup"><span data-stu-id="67946-1051">1.0</span></span>|
|[<span data-ttu-id="67946-1052">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="67946-1052">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="67946-1053">ReadItem</span><span class="sxs-lookup"><span data-stu-id="67946-1053">ReadItem</span></span>|
|[<span data-ttu-id="67946-1054">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="67946-1054">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="67946-1055">阅读</span><span class="sxs-lookup"><span data-stu-id="67946-1055">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="67946-1056">返回：</span><span class="sxs-lookup"><span data-stu-id="67946-1056">Returns:</span></span>

<span data-ttu-id="67946-p160">如果具有匹配 `name` 参数的 `FilterName` 元素值的清单中没有任何 `ItemHasKnownEntity` 元素，则该方法返回 `null`。如果 `name` 参数匹配清单中的 `ItemHasKnownEntity` 元素，但在匹配的当前项目中没有实体，则该方法返回一个空数组。</span><span class="sxs-lookup"><span data-stu-id="67946-p160">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="67946-1059">类型：Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="67946-1059">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

#### <a name="getinitializationcontextasyncoptions-callback"></a><span data-ttu-id="67946-1060">getInitializationContextAsync([options], [callback])</span><span class="sxs-lookup"><span data-stu-id="67946-1060">getInitializationContextAsync([options], [callback])</span></span>

<span data-ttu-id="67946-1061">当[可操作邮件激活](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message)加载项时，获取传递的初始化数据。</span><span class="sxs-lookup"><span data-stu-id="67946-1061">Gets initialization data passed when the add-in is [activated by an actionable message](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

> [!NOTE]
> <span data-ttu-id="67946-1062">仅支持使用此方法通过 Outlook 2016 Windows （单击即点即用版本大于 16.0.8413.1000） 和 Outlook web 上的 Office 365。</span><span class="sxs-lookup"><span data-stu-id="67946-1062">This method is only supported by Outlook 2016 for Windows (Click-to-Run versions greater than 16.0.8413.1000) and Outlook on the web for Office 365.</span></span>

##### <a name="parameters"></a><span data-ttu-id="67946-1063">参数：</span><span class="sxs-lookup"><span data-stu-id="67946-1063">Parameters:</span></span>
|<span data-ttu-id="67946-1064">名称</span><span class="sxs-lookup"><span data-stu-id="67946-1064">Name</span></span>|<span data-ttu-id="67946-1065">类型</span><span class="sxs-lookup"><span data-stu-id="67946-1065">Type</span></span>|<span data-ttu-id="67946-1066">属性</span><span class="sxs-lookup"><span data-stu-id="67946-1066">Attributes</span></span>|<span data-ttu-id="67946-1067">说明</span><span class="sxs-lookup"><span data-stu-id="67946-1067">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="67946-1068">对象</span><span class="sxs-lookup"><span data-stu-id="67946-1068">Object</span></span>|<span data-ttu-id="67946-1069">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="67946-1069">&lt;optional&gt;</span></span>|<span data-ttu-id="67946-1070">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="67946-1070">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="67946-1071">对象</span><span class="sxs-lookup"><span data-stu-id="67946-1071">Object</span></span>|<span data-ttu-id="67946-1072">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="67946-1072">&lt;optional&gt;</span></span>|<span data-ttu-id="67946-1073">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="67946-1073">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="67946-1074">函数</span><span class="sxs-lookup"><span data-stu-id="67946-1074">function</span></span>|<span data-ttu-id="67946-1075">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="67946-1075">&lt;optional&gt;</span></span>|<span data-ttu-id="67946-1076">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="67946-1076">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="67946-1077">成功，在初始化数据提供在`asyncResult.value`属性的字符串形式。</span><span class="sxs-lookup"><span data-stu-id="67946-1077">On success, the initialization data is provided in the `asyncResult.value` property as a string.</span></span><br/><span data-ttu-id="67946-1078">如果没有初始化上下文，`asyncResult` 对象包含 `Error` 对象，并将它的 `code` 和 `name` 属性分别设置为 `9020` 和 `GenericResponseError`。</span><span class="sxs-lookup"><span data-stu-id="67946-1078">If there is no initialization context, the `asyncResult` object will contain an `Error` object with its `code` property set to `9020` and its `name` property set to `GenericResponseError`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="67946-1079">要求</span><span class="sxs-lookup"><span data-stu-id="67946-1079">Requirements</span></span>

|<span data-ttu-id="67946-1080">要求</span><span class="sxs-lookup"><span data-stu-id="67946-1080">Requirement</span></span>|<span data-ttu-id="67946-1081">值</span><span class="sxs-lookup"><span data-stu-id="67946-1081">Value</span></span>|
|---|---|
|[<span data-ttu-id="67946-1082">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="67946-1082">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="67946-1083">预览</span><span class="sxs-lookup"><span data-stu-id="67946-1083">Preview</span></span>|
|[<span data-ttu-id="67946-1084">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="67946-1084">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="67946-1085">ReadItem</span><span class="sxs-lookup"><span data-stu-id="67946-1085">ReadItem</span></span>|
|[<span data-ttu-id="67946-1086">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="67946-1086">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="67946-1087">阅读</span><span class="sxs-lookup"><span data-stu-id="67946-1087">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="67946-1088">示例</span><span class="sxs-lookup"><span data-stu-id="67946-1088">Example</span></span>

```
// Get the initialization context (if present)
Office.context.mailbox.item.getInitializationContextAsync(
  function(asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
      if (asyncResult.value != null && asyncResult.value.length > 0) {
        // The value is a string, parse to an object
        var context = JSON.parse(asyncResult.value);
        // Do something with context
      } else {
        // Empty context, treat as no context
      }
    } else {
      if (asyncResult.error.code == 9020) {
        // GenericResponseError returned when there is
        // no context
        // Treat as no context
      } else {
        // Handle the error
      }
    }
  }
);
```

#### <a name="getregexmatches--object"></a><span data-ttu-id="67946-1089">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="67946-1089">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="67946-1090">返回所选项目中匹配在清单 XML 文件中定义的正则表达式的字符串值。</span><span class="sxs-lookup"><span data-stu-id="67946-1090">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="67946-1091">在适用于 iOS 的 Outlook 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="67946-1091">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="67946-p161">`getRegExMatches` 方法返回匹配在清单 XML 文件中的每个 `ItemHasRegularExpressionMatch` 或 `ItemHasKnownEntity` 规则元素中定义的正则表达式的字符串。对于 `ItemHasRegularExpressionMatch` 规则，匹配字符串必须发生在该规则指定的项目的属性中。`PropertyName` 简单类型定义支持的属性。</span><span class="sxs-lookup"><span data-stu-id="67946-p161">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="67946-1095">例如，考虑一个外接程序清单具有以下 `Rule` 元素：</span><span class="sxs-lookup"><span data-stu-id="67946-1095">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="67946-1096">从 `getRegExMatches` 返回的对象应有两个属性：`fruits` 和 `veggies`。</span><span class="sxs-lookup"><span data-stu-id="67946-1096">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="67946-p162">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。而是使用 [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) 方法检索整个正文。</span><span class="sxs-lookup"><span data-stu-id="67946-p162">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="67946-1100">Requirements</span><span class="sxs-lookup"><span data-stu-id="67946-1100">Requirements</span></span>

|<span data-ttu-id="67946-1101">要求</span><span class="sxs-lookup"><span data-stu-id="67946-1101">Requirement</span></span>|<span data-ttu-id="67946-1102">值</span><span class="sxs-lookup"><span data-stu-id="67946-1102">Value</span></span>|
|---|---|
|[<span data-ttu-id="67946-1103">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="67946-1103">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="67946-1104">1.0</span><span class="sxs-lookup"><span data-stu-id="67946-1104">1.0</span></span>|
|[<span data-ttu-id="67946-1105">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="67946-1105">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="67946-1106">ReadItem</span><span class="sxs-lookup"><span data-stu-id="67946-1106">ReadItem</span></span>|
|[<span data-ttu-id="67946-1107">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="67946-1107">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="67946-1108">阅读</span><span class="sxs-lookup"><span data-stu-id="67946-1108">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="67946-1109">返回：</span><span class="sxs-lookup"><span data-stu-id="67946-1109">Returns:</span></span>

<span data-ttu-id="67946-p163">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串数组的对象。每个数组的名称等于匹配 `ItemHasRegularExpressionMatch` 规则的 `RegExName` 属性或匹配 `ItemHasKnownEntity` 规则的 `FilterName` 属性的相应值。</span><span class="sxs-lookup"><span data-stu-id="67946-p163">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="67946-1112">

<dt>
类型</dt>


</span><span class="sxs-lookup"><span data-stu-id="67946-1112">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="67946-1113">对象</span><span class="sxs-lookup"><span data-stu-id="67946-1113">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="67946-1114">示例</span><span class="sxs-lookup"><span data-stu-id="67946-1114">Example</span></span>

<span data-ttu-id="67946-1115">以下示例显示了如何访问正则表达式规则元素 `fruits` 和 `veggies` 的匹配项的数组，这些元素在清单中指定。</span><span class="sxs-lookup"><span data-stu-id="67946-1115">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="67946-1116">getregexmatchesbyname （name） → (nullable) {数组。 < 字符串 >}</span><span class="sxs-lookup"><span data-stu-id="67946-1116">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="67946-1117">返回匹配在清单 XML 文件中定义的命名正则表达式的所选项目中的字符串值。</span><span class="sxs-lookup"><span data-stu-id="67946-1117">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="67946-1118">在适用于 iOS 的 Outlook 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="67946-1118">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="67946-1119">`getRegExMatchesByName` 方法返回匹配在具有指定 `RegExName` 元素值的清单 XML 文件中的 `ItemHasRegularExpressionMatch` 规则元素中定义的正则表达式的字符串。</span><span class="sxs-lookup"><span data-stu-id="67946-1119">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="67946-p164">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。</span><span class="sxs-lookup"><span data-stu-id="67946-p164">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="67946-1122">参数：</span><span class="sxs-lookup"><span data-stu-id="67946-1122">Parameters:</span></span>

|<span data-ttu-id="67946-1123">名称</span><span class="sxs-lookup"><span data-stu-id="67946-1123">Name</span></span>|<span data-ttu-id="67946-1124">类型</span><span class="sxs-lookup"><span data-stu-id="67946-1124">Type</span></span>|<span data-ttu-id="67946-1125">说明</span><span class="sxs-lookup"><span data-stu-id="67946-1125">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="67946-1126">字符串</span><span class="sxs-lookup"><span data-stu-id="67946-1126">String</span></span>|<span data-ttu-id="67946-1127">定义筛选器匹配的 `ItemHasRegularExpressionMatch` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="67946-1127">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="67946-1128">要求</span><span class="sxs-lookup"><span data-stu-id="67946-1128">Requirements</span></span>

|<span data-ttu-id="67946-1129">要求</span><span class="sxs-lookup"><span data-stu-id="67946-1129">Requirement</span></span>|<span data-ttu-id="67946-1130">值</span><span class="sxs-lookup"><span data-stu-id="67946-1130">Value</span></span>|
|---|---|
|[<span data-ttu-id="67946-1131">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="67946-1131">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="67946-1132">1.0</span><span class="sxs-lookup"><span data-stu-id="67946-1132">1.0</span></span>|
|[<span data-ttu-id="67946-1133">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="67946-1133">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="67946-1134">ReadItem</span><span class="sxs-lookup"><span data-stu-id="67946-1134">ReadItem</span></span>|
|[<span data-ttu-id="67946-1135">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="67946-1135">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="67946-1136">阅读</span><span class="sxs-lookup"><span data-stu-id="67946-1136">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="67946-1137">返回：</span><span class="sxs-lookup"><span data-stu-id="67946-1137">Returns:</span></span>

<span data-ttu-id="67946-1138">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串的数组。</span><span class="sxs-lookup"><span data-stu-id="67946-1138">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="67946-1139">

<dt>类型</dt>

</span><span class="sxs-lookup"><span data-stu-id="67946-1139">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="67946-1140">数组。 < 字符串 ></span><span class="sxs-lookup"><span data-stu-id="67946-1140">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="67946-1141">示例</span><span class="sxs-lookup"><span data-stu-id="67946-1141">Example</span></span>

```
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="67946-1142">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="67946-1142">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="67946-1143">以异步方式返回邮件的主题或正文中选定的数据。</span><span class="sxs-lookup"><span data-stu-id="67946-1143">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="67946-p165">如果没有选定内容，但光标位于正文或主题中，此方法将会为所选数据返回 null。如果选定的是字段，而不是正文或主题，则此方法返回 `InvalidSelection` 错误。</span><span class="sxs-lookup"><span data-stu-id="67946-p165">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="67946-1146">参数：</span><span class="sxs-lookup"><span data-stu-id="67946-1146">Parameters:</span></span>

|<span data-ttu-id="67946-1147">名称</span><span class="sxs-lookup"><span data-stu-id="67946-1147">Name</span></span>|<span data-ttu-id="67946-1148">类型</span><span class="sxs-lookup"><span data-stu-id="67946-1148">Type</span></span>|<span data-ttu-id="67946-1149">属性</span><span class="sxs-lookup"><span data-stu-id="67946-1149">Attributes</span></span>|<span data-ttu-id="67946-1150">说明</span><span class="sxs-lookup"><span data-stu-id="67946-1150">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="67946-1151">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="67946-1151">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="67946-p166">请求数据的格式。如果为文本，则此方法返回纯文本作为字符串，删除任何显示的 HTML 标记。如果为 HTML，则此方法返回所选文本，不论是纯文本还是 HTML。</span><span class="sxs-lookup"><span data-stu-id="67946-p166">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="67946-1155">对象</span><span class="sxs-lookup"><span data-stu-id="67946-1155">Object</span></span>|<span data-ttu-id="67946-1156">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="67946-1156">&lt;optional&gt;</span></span>|<span data-ttu-id="67946-1157">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="67946-1157">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="67946-1158">对象</span><span class="sxs-lookup"><span data-stu-id="67946-1158">Object</span></span>|<span data-ttu-id="67946-1159">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="67946-1159">&lt;optional&gt;</span></span>|<span data-ttu-id="67946-1160">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="67946-1160">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="67946-1161">函数</span><span class="sxs-lookup"><span data-stu-id="67946-1161">function</span></span>||<span data-ttu-id="67946-1162">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="67946-1162">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="67946-1163">若要从回调方法访问所选数据，请调用 `asyncResult.value.data`。</span><span class="sxs-lookup"><span data-stu-id="67946-1163">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="67946-1164">若要访问所选内容来自源属性，请调用`asyncResult.value.sourceProperty`，其中将为`body`或`subject`。</span><span class="sxs-lookup"><span data-stu-id="67946-1164">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="67946-1165">要求</span><span class="sxs-lookup"><span data-stu-id="67946-1165">Requirements</span></span>

|<span data-ttu-id="67946-1166">要求</span><span class="sxs-lookup"><span data-stu-id="67946-1166">Requirement</span></span>|<span data-ttu-id="67946-1167">值</span><span class="sxs-lookup"><span data-stu-id="67946-1167">Value</span></span>|
|---|---|
|[<span data-ttu-id="67946-1168">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="67946-1168">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="67946-1169">1.2</span><span class="sxs-lookup"><span data-stu-id="67946-1169">1.2</span></span>|
|[<span data-ttu-id="67946-1170">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="67946-1170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="67946-1171">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="67946-1171">ReadWriteItem</span></span>|
|[<span data-ttu-id="67946-1172">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="67946-1172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="67946-1173">撰写</span><span class="sxs-lookup"><span data-stu-id="67946-1173">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="67946-1174">返回：</span><span class="sxs-lookup"><span data-stu-id="67946-1174">Returns:</span></span>

<span data-ttu-id="67946-1175">作为字符串的所选数据的格式由 `coercionType` 确定。</span><span class="sxs-lookup"><span data-stu-id="67946-1175">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="67946-1176">

<dt>
类型</dt>


</span><span class="sxs-lookup"><span data-stu-id="67946-1176">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="67946-1177">字符串</span><span class="sxs-lookup"><span data-stu-id="67946-1177">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="67946-1178">示例</span><span class="sxs-lookup"><span data-stu-id="67946-1178">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="67946-1179">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="67946-1179">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="67946-p168">获取在用户已选择的突出显示匹配项中找到的实体。突出显示匹配项适用于[上下文外接程序](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins)。</span><span class="sxs-lookup"><span data-stu-id="67946-p168">Gets the entities found in a highlighted match a user has selected. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="67946-1182">在适用于 iOS 的 Outlook 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="67946-1182">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="67946-1183">要求</span><span class="sxs-lookup"><span data-stu-id="67946-1183">Requirements</span></span>

|<span data-ttu-id="67946-1184">要求</span><span class="sxs-lookup"><span data-stu-id="67946-1184">Requirement</span></span>|<span data-ttu-id="67946-1185">值</span><span class="sxs-lookup"><span data-stu-id="67946-1185">Value</span></span>|
|---|---|
|[<span data-ttu-id="67946-1186">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="67946-1186">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="67946-1187">1.6</span><span class="sxs-lookup"><span data-stu-id="67946-1187">1.6</span></span>|
|[<span data-ttu-id="67946-1188">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="67946-1188">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="67946-1189">ReadItem</span><span class="sxs-lookup"><span data-stu-id="67946-1189">ReadItem</span></span>|
|[<span data-ttu-id="67946-1190">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="67946-1190">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="67946-1191">阅读</span><span class="sxs-lookup"><span data-stu-id="67946-1191">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="67946-1192">返回：</span><span class="sxs-lookup"><span data-stu-id="67946-1192">Returns:</span></span>

<span data-ttu-id="67946-1193">类型：[Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="67946-1193">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="67946-1194">示例</span><span class="sxs-lookup"><span data-stu-id="67946-1194">Example</span></span>

<span data-ttu-id="67946-1195">以下示例访问用户选择的突出显示匹配项中的地址实体。</span><span class="sxs-lookup"><span data-stu-id="67946-1195">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="67946-1196">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="67946-1196">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="67946-p169">返回突出显示匹配项中匹配在清单 XML 文件中定义的正则表达式的字符串值。突出显示匹配项适用于[上下文外接程序](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins)。</span><span class="sxs-lookup"><span data-stu-id="67946-p169">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="67946-1199">在适用于 iOS 的 Outlook 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="67946-1199">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="67946-p170">`getSelectedRegExMatches` 方法返回匹配在清单 XML 文件中的每个 `ItemHasRegularExpressionMatch` 或 `ItemHasKnownEntity` 规则元素中定义的正则表达式的字符串。对于 `ItemHasRegularExpressionMatch` 规则，匹配字符串必须发生在该规则指定的项目的属性中。`PropertyName` 简单类型定义支持的属性。</span><span class="sxs-lookup"><span data-stu-id="67946-p170">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="67946-1203">例如，考虑一个外接程序清单具有以下 `Rule` 元素：</span><span class="sxs-lookup"><span data-stu-id="67946-1203">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="67946-1204">从 `getRegExMatches` 返回的对象应有两个属性：`fruits` 和 `veggies`。</span><span class="sxs-lookup"><span data-stu-id="67946-1204">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="67946-p171">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。而是使用 [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) 方法检索整个正文。</span><span class="sxs-lookup"><span data-stu-id="67946-p171">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="67946-1208">Requirements</span><span class="sxs-lookup"><span data-stu-id="67946-1208">Requirements</span></span>

|<span data-ttu-id="67946-1209">要求</span><span class="sxs-lookup"><span data-stu-id="67946-1209">Requirement</span></span>|<span data-ttu-id="67946-1210">值</span><span class="sxs-lookup"><span data-stu-id="67946-1210">Value</span></span>|
|---|---|
|[<span data-ttu-id="67946-1211">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="67946-1211">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="67946-1212">1.6</span><span class="sxs-lookup"><span data-stu-id="67946-1212">1.6</span></span>|
|[<span data-ttu-id="67946-1213">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="67946-1213">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="67946-1214">ReadItem</span><span class="sxs-lookup"><span data-stu-id="67946-1214">ReadItem</span></span>|
|[<span data-ttu-id="67946-1215">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="67946-1215">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="67946-1216">阅读</span><span class="sxs-lookup"><span data-stu-id="67946-1216">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="67946-1217">返回：</span><span class="sxs-lookup"><span data-stu-id="67946-1217">Returns:</span></span>

<span data-ttu-id="67946-p172">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串数组的对象。每个数组的名称等于匹配 `ItemHasRegularExpressionMatch` 规则的 `RegExName` 属性或匹配 `ItemHasKnownEntity` 规则的 `FilterName` 属性的相应值。</span><span class="sxs-lookup"><span data-stu-id="67946-p172">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="67946-1220">示例</span><span class="sxs-lookup"><span data-stu-id="67946-1220">Example</span></span>

<span data-ttu-id="67946-1221">以下示例显示了如何访问正则表达式规则元素 `fruits` 和 `veggies` 的匹配项的数组，这些元素在清单中指定。</span><span class="sxs-lookup"><span data-stu-id="67946-1221">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="67946-1222">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="67946-1222">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="67946-1223">异步加载所选项目上此外接程序的自定义属性。</span><span class="sxs-lookup"><span data-stu-id="67946-1223">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="67946-p173">自定义属性基于每个应用、每个项目存储为键/值对。此方法在回调中返回 `CustomProperties` 对象，该回调提供访问特定于当前项目和当前外接程序的自定义属性的方法。自定义属性未在项目上加密，因此这不应用作安全存储。</span><span class="sxs-lookup"><span data-stu-id="67946-p173">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="67946-1227">参数：</span><span class="sxs-lookup"><span data-stu-id="67946-1227">Parameters:</span></span>

|<span data-ttu-id="67946-1228">名称</span><span class="sxs-lookup"><span data-stu-id="67946-1228">Name</span></span>|<span data-ttu-id="67946-1229">类型</span><span class="sxs-lookup"><span data-stu-id="67946-1229">Type</span></span>|<span data-ttu-id="67946-1230">属性</span><span class="sxs-lookup"><span data-stu-id="67946-1230">Attributes</span></span>|<span data-ttu-id="67946-1231">说明</span><span class="sxs-lookup"><span data-stu-id="67946-1231">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="67946-1232">函数</span><span class="sxs-lookup"><span data-stu-id="67946-1232">function</span></span>||<span data-ttu-id="67946-1233">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="67946-1233">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="67946-1234">自定义属性作为 `asyncResult.value` 属性中的 [`CustomProperties`](/javascript/api/outlook/office.customproperties) 对象提供。</span><span class="sxs-lookup"><span data-stu-id="67946-1234">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="67946-1235">此对象可用于获取、 设置，和从项目中删除自定义属性并将更改保存到自定义属性设置回服务器。</span><span class="sxs-lookup"><span data-stu-id="67946-1235">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="67946-1236">对象</span><span class="sxs-lookup"><span data-stu-id="67946-1236">Object</span></span>|<span data-ttu-id="67946-1237">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="67946-1237">&lt;optional&gt;</span></span>|<span data-ttu-id="67946-1238">开发人员可以提供他们想要访问的回调函数中的任何对象。</span><span class="sxs-lookup"><span data-stu-id="67946-1238">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="67946-1239">此对象可以访问`asyncResult.asyncContext`的回调函数中的属性。</span><span class="sxs-lookup"><span data-stu-id="67946-1239">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="67946-1240">要求</span><span class="sxs-lookup"><span data-stu-id="67946-1240">Requirements</span></span>

|<span data-ttu-id="67946-1241">要求</span><span class="sxs-lookup"><span data-stu-id="67946-1241">Requirement</span></span>|<span data-ttu-id="67946-1242">值</span><span class="sxs-lookup"><span data-stu-id="67946-1242">Value</span></span>|
|---|---|
|[<span data-ttu-id="67946-1243">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="67946-1243">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="67946-1244">1.0</span><span class="sxs-lookup"><span data-stu-id="67946-1244">1.0</span></span>|
|[<span data-ttu-id="67946-1245">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="67946-1245">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="67946-1246">ReadItem</span><span class="sxs-lookup"><span data-stu-id="67946-1246">ReadItem</span></span>|
|[<span data-ttu-id="67946-1247">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="67946-1247">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="67946-1248">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="67946-1248">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="67946-1249">示例</span><span class="sxs-lookup"><span data-stu-id="67946-1249">Example</span></span>

<span data-ttu-id="67946-p176">以下代码示例显示了如何使用 `loadCustomPropertiesAsync` 方法异步加载特定于当前项目的自定义属性。该示例还显示了如何使用 `CustomProperties.saveAsync` 方法将这些属性重新保存到服务器。加载自定义属性后，该代码示例将使用 `CustomProperties.get` 方法读取自定义属性 `myProp`，使用 `CustomProperties.set` 方法写入自定义属性 `otherProp`，最后调用 `saveAsync` 方法保存这些自定义属性。</span><span class="sxs-lookup"><span data-stu-id="67946-p176">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="67946-1253">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="67946-1253">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="67946-1254">将附件从邮件或约会中删除。</span><span class="sxs-lookup"><span data-stu-id="67946-1254">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="67946-p177">`removeAttachmentAsync` 方法删除项目中带指定标识符的附件。最佳做法是，仅当同一个邮件应用程序在同一会话中添加了一个附件时，你才应使用该附件标识符来删除该附件。在 Outlook Web App 和适用于设备的 OWA 中，附件标识符只在同一个会话中才有效。当用户关闭应用程序，或者如果用户开始在内嵌窗体中撰写，并在随后弹出的内嵌窗体中继续在单独的窗口撰写时，会话即结束。</span><span class="sxs-lookup"><span data-stu-id="67946-p177">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="67946-1259">参数：</span><span class="sxs-lookup"><span data-stu-id="67946-1259">Parameters:</span></span>

|<span data-ttu-id="67946-1260">名称</span><span class="sxs-lookup"><span data-stu-id="67946-1260">Name</span></span>|<span data-ttu-id="67946-1261">类型</span><span class="sxs-lookup"><span data-stu-id="67946-1261">Type</span></span>|<span data-ttu-id="67946-1262">属性</span><span class="sxs-lookup"><span data-stu-id="67946-1262">Attributes</span></span>|<span data-ttu-id="67946-1263">说明</span><span class="sxs-lookup"><span data-stu-id="67946-1263">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="67946-1264">字符串</span><span class="sxs-lookup"><span data-stu-id="67946-1264">String</span></span>||<span data-ttu-id="67946-p178">要删除的附件的标识符。字符串的最大长度为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="67946-p178">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`|<span data-ttu-id="67946-1267">对象</span><span class="sxs-lookup"><span data-stu-id="67946-1267">Object</span></span>|<span data-ttu-id="67946-1268">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="67946-1268">&lt;optional&gt;</span></span>|<span data-ttu-id="67946-1269">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="67946-1269">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="67946-1270">对象</span><span class="sxs-lookup"><span data-stu-id="67946-1270">Object</span></span>|<span data-ttu-id="67946-1271">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="67946-1271">&lt;optional&gt;</span></span>|<span data-ttu-id="67946-1272">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="67946-1272">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="67946-1273">函数</span><span class="sxs-lookup"><span data-stu-id="67946-1273">function</span></span>|<span data-ttu-id="67946-1274">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="67946-1274">&lt;optional&gt;</span></span>|<span data-ttu-id="67946-1275">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="67946-1275">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="67946-1276">如果删除附件失败，`asyncResult.error` 属性将包含一个说明失败原因的错误代码。</span><span class="sxs-lookup"><span data-stu-id="67946-1276">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="67946-1277">错误</span><span class="sxs-lookup"><span data-stu-id="67946-1277">Errors</span></span>

|<span data-ttu-id="67946-1278">错误代码</span><span class="sxs-lookup"><span data-stu-id="67946-1278">Error code</span></span>|<span data-ttu-id="67946-1279">说明</span><span class="sxs-lookup"><span data-stu-id="67946-1279">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="67946-1280">附件标识符不存在。</span><span class="sxs-lookup"><span data-stu-id="67946-1280">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="67946-1281">要求</span><span class="sxs-lookup"><span data-stu-id="67946-1281">Requirements</span></span>

|<span data-ttu-id="67946-1282">要求</span><span class="sxs-lookup"><span data-stu-id="67946-1282">Requirement</span></span>|<span data-ttu-id="67946-1283">值</span><span class="sxs-lookup"><span data-stu-id="67946-1283">Value</span></span>|
|---|---|
|[<span data-ttu-id="67946-1284">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="67946-1284">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="67946-1285">1.1</span><span class="sxs-lookup"><span data-stu-id="67946-1285">1.1</span></span>|
|[<span data-ttu-id="67946-1286">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="67946-1286">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="67946-1287">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="67946-1287">ReadWriteItem</span></span>|
|[<span data-ttu-id="67946-1288">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="67946-1288">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="67946-1289">撰写</span><span class="sxs-lookup"><span data-stu-id="67946-1289">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="67946-1290">示例</span><span class="sxs-lookup"><span data-stu-id="67946-1290">Example</span></span>

<span data-ttu-id="67946-1291">以下代码删除包含标识符 '0' 的附件。</span><span class="sxs-lookup"><span data-stu-id="67946-1291">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="removehandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="67946-1292">removeHandlerAsync (eventType，处理程序中，[选项] [回调])</span><span class="sxs-lookup"><span data-stu-id="67946-1292">removeHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="67946-1293">删除事件处理程序支持的事件。</span><span class="sxs-lookup"><span data-stu-id="67946-1293">Removes an event handler for a supported event.</span></span>

<span data-ttu-id="67946-1294">当前受支持的事件类型的`Office.EventType.AppointmentTimeChanged`， `Office.EventType.RecipientsChanged`，和`Office.EventType.RecurrencePatternChanged`</span><span class="sxs-lookup"><span data-stu-id="67946-1294">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrencePatternChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="67946-1295">参数：</span><span class="sxs-lookup"><span data-stu-id="67946-1295">Parameters:</span></span>

| <span data-ttu-id="67946-1296">名称</span><span class="sxs-lookup"><span data-stu-id="67946-1296">Name</span></span> | <span data-ttu-id="67946-1297">类型</span><span class="sxs-lookup"><span data-stu-id="67946-1297">Type</span></span> | <span data-ttu-id="67946-1298">属性</span><span class="sxs-lookup"><span data-stu-id="67946-1298">Attributes</span></span> | <span data-ttu-id="67946-1299">说明</span><span class="sxs-lookup"><span data-stu-id="67946-1299">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="67946-1300">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="67946-1300">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="67946-1301">应调用处理程序的事件。</span><span class="sxs-lookup"><span data-stu-id="67946-1301">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="67946-1302">函数</span><span class="sxs-lookup"><span data-stu-id="67946-1302">Function</span></span> || <span data-ttu-id="67946-p179">用于处理事件的函数。此函数必须接受一个参数，即对象文本。参数上的 `type` 属性将匹配传递给 `removeHandlerAsync` 的 `eventType` 参数。</span><span class="sxs-lookup"><span data-stu-id="67946-p179">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `removeHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="67946-1306">Object</span><span class="sxs-lookup"><span data-stu-id="67946-1306">Object</span></span> | <span data-ttu-id="67946-1307">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="67946-1307">&lt;optional&gt;</span></span> | <span data-ttu-id="67946-1308">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="67946-1308">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="67946-1309">对象</span><span class="sxs-lookup"><span data-stu-id="67946-1309">Object</span></span> | <span data-ttu-id="67946-1310">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="67946-1310">&lt;optional&gt;</span></span> | <span data-ttu-id="67946-1311">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="67946-1311">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="67946-1312">函数</span><span class="sxs-lookup"><span data-stu-id="67946-1312">function</span></span>| <span data-ttu-id="67946-1313">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="67946-1313">&lt;optional&gt;</span></span>|<span data-ttu-id="67946-1314">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="67946-1314">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="67946-1315">Requirements</span><span class="sxs-lookup"><span data-stu-id="67946-1315">Requirements</span></span>

|<span data-ttu-id="67946-1316">要求</span><span class="sxs-lookup"><span data-stu-id="67946-1316">Requirement</span></span>| <span data-ttu-id="67946-1317">值</span><span class="sxs-lookup"><span data-stu-id="67946-1317">Value</span></span>|
|---|---|
|[<span data-ttu-id="67946-1318">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="67946-1318">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="67946-1319">预览</span><span class="sxs-lookup"><span data-stu-id="67946-1319">Preview</span></span> |
|[<span data-ttu-id="67946-1320">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="67946-1320">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="67946-1321">ReadItem</span><span class="sxs-lookup"><span data-stu-id="67946-1321">ReadItem</span></span> |
|[<span data-ttu-id="67946-1322">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="67946-1322">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="67946-1323">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="67946-1323">Compose or read</span></span> |

##### <a name="example"></a><span data-ttu-id="67946-1324">示例</span><span class="sxs-lookup"><span data-stu-id="67946-1324">Example</span></span>

```
Office.initialize = function (reason) {
  $(document).ready(function () {
    Office.context.mailbox.item.removeHandlerAsync(Office.EventType.RecurrencePatternChanged, loadNewItem, function (result) {
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

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="67946-1325">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="67946-1325">saveAsync([options], callback)</span></span>

<span data-ttu-id="67946-1326">异步保存项目。</span><span class="sxs-lookup"><span data-stu-id="67946-1326">Asynchronously saves an item.</span></span>

<span data-ttu-id="67946-p180">调用时，此方法将当前邮件保存为草稿，并通过回调方法返回项目 ID。在 Outlook Web App 或 Outlook 联机模式下，该项目被保存到服务器中。在 Outlook 缓存模式下，该项目被保存到本地缓存中。</span><span class="sxs-lookup"><span data-stu-id="67946-p180">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="67946-1330">如果加载项调用`saveAsync`中的项目在撰写模式下才能获取`itemId`若要使用 EWS 或 REST API，请注意，缓存模式 Outlook 时，可能需要一些时间才能项目实际同步到服务器。</span><span class="sxs-lookup"><span data-stu-id="67946-1330">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="67946-1331">直到该项目同步，使用`itemId`将返回错误。</span><span class="sxs-lookup"><span data-stu-id="67946-1331">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="67946-p182">由于约会没有草稿状态，如果以撰写模式在约会中调用 `saveAsync`，则该项将被保存为用户日历中的正常约会。对于之前未保存过的新约会，则不会发送邀请。保存现有约会将向添加或删除的与会者发送更新。</span><span class="sxs-lookup"><span data-stu-id="67946-p182">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="67946-1335">以下客户端具有不同行为的`saveAsync`上约会的撰写模式下：</span><span class="sxs-lookup"><span data-stu-id="67946-1335">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="67946-1336">Mac Outlook 不支持`saveAsync`中的会议在撰写模式。</span><span class="sxs-lookup"><span data-stu-id="67946-1336">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="67946-1337">调用`saveAsync`上 Mac Outlook 中的会议，则将返回错误。</span><span class="sxs-lookup"><span data-stu-id="67946-1337">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="67946-1338">在 web 上的 outlook 始终发送的邀请或更新何时`saveAsync`上约会调用在撰写模式。</span><span class="sxs-lookup"><span data-stu-id="67946-1338">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="67946-1339">参数：</span><span class="sxs-lookup"><span data-stu-id="67946-1339">Parameters:</span></span>

|<span data-ttu-id="67946-1340">名称</span><span class="sxs-lookup"><span data-stu-id="67946-1340">Name</span></span>|<span data-ttu-id="67946-1341">类型</span><span class="sxs-lookup"><span data-stu-id="67946-1341">Type</span></span>|<span data-ttu-id="67946-1342">属性</span><span class="sxs-lookup"><span data-stu-id="67946-1342">Attributes</span></span>|<span data-ttu-id="67946-1343">说明</span><span class="sxs-lookup"><span data-stu-id="67946-1343">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="67946-1344">对象</span><span class="sxs-lookup"><span data-stu-id="67946-1344">Object</span></span>|<span data-ttu-id="67946-1345">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="67946-1345">&lt;optional&gt;</span></span>|<span data-ttu-id="67946-1346">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="67946-1346">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="67946-1347">对象</span><span class="sxs-lookup"><span data-stu-id="67946-1347">Object</span></span>|<span data-ttu-id="67946-1348">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="67946-1348">&lt;optional&gt;</span></span>|<span data-ttu-id="67946-1349">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="67946-1349">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="67946-1350">函数</span><span class="sxs-lookup"><span data-stu-id="67946-1350">function</span></span>||<span data-ttu-id="67946-1351">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="67946-1351">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="67946-1352">中提供的项标识符是成功，`asyncResult.value`属性。</span><span class="sxs-lookup"><span data-stu-id="67946-1352">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="67946-1353">要求</span><span class="sxs-lookup"><span data-stu-id="67946-1353">Requirements</span></span>

|<span data-ttu-id="67946-1354">要求</span><span class="sxs-lookup"><span data-stu-id="67946-1354">Requirement</span></span>|<span data-ttu-id="67946-1355">值</span><span class="sxs-lookup"><span data-stu-id="67946-1355">Value</span></span>|
|---|---|
|[<span data-ttu-id="67946-1356">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="67946-1356">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="67946-1357">1.3</span><span class="sxs-lookup"><span data-stu-id="67946-1357">1.3</span></span>|
|[<span data-ttu-id="67946-1358">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="67946-1358">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="67946-1359">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="67946-1359">ReadWriteItem</span></span>|
|[<span data-ttu-id="67946-1360">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="67946-1360">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="67946-1361">撰写</span><span class="sxs-lookup"><span data-stu-id="67946-1361">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="67946-1362">示例</span><span class="sxs-lookup"><span data-stu-id="67946-1362">Examples</span></span>

```
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="67946-p184">下面是传递给回调函数的 `result` 参数的示例。`value` 属性包含的项目的项目 ID。</span><span class="sxs-lookup"><span data-stu-id="67946-p184">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="67946-1365">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="67946-1365">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="67946-1366">以异步方式将数据插入到邮件的正文或主题中。</span><span class="sxs-lookup"><span data-stu-id="67946-1366">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="67946-p185">`setSelectedDataAsync` 方法将指定字符串插入到项目主题或正文的光标位置，或者，如果在编辑器中已选择文本，则该方法将替换选择的文本。如果光标不在正文或主题字段中，则返回错误。插入之后，光标会位于插入内容的末尾。</span><span class="sxs-lookup"><span data-stu-id="67946-p185">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="67946-1370">参数：</span><span class="sxs-lookup"><span data-stu-id="67946-1370">Parameters:</span></span>

|<span data-ttu-id="67946-1371">名称</span><span class="sxs-lookup"><span data-stu-id="67946-1371">Name</span></span>|<span data-ttu-id="67946-1372">类型</span><span class="sxs-lookup"><span data-stu-id="67946-1372">Type</span></span>|<span data-ttu-id="67946-1373">属性</span><span class="sxs-lookup"><span data-stu-id="67946-1373">Attributes</span></span>|<span data-ttu-id="67946-1374">说明</span><span class="sxs-lookup"><span data-stu-id="67946-1374">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="67946-1375">字符串</span><span class="sxs-lookup"><span data-stu-id="67946-1375">String</span></span>||<span data-ttu-id="67946-p186">要插入的数据。数据不得超过 1,000,000 个字符。如果传入的数据超过 1,000,000 个字符，则会引发 `ArgumentOutOfRange` 异常。</span><span class="sxs-lookup"><span data-stu-id="67946-p186">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="67946-1379">Object</span><span class="sxs-lookup"><span data-stu-id="67946-1379">Object</span></span>|<span data-ttu-id="67946-1380">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="67946-1380">&lt;optional&gt;</span></span>|<span data-ttu-id="67946-1381">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="67946-1381">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="67946-1382">对象</span><span class="sxs-lookup"><span data-stu-id="67946-1382">Object</span></span>|<span data-ttu-id="67946-1383">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="67946-1383">&lt;optional&gt;</span></span>|<span data-ttu-id="67946-1384">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="67946-1384">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="67946-1385">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="67946-1385">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="67946-1386">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="67946-1386">&lt;optional&gt;</span></span>|<span data-ttu-id="67946-p187">如果为 `text`，则在 Outlook Web App 和 Outlook 中应用当前样式。如果该字段是 HTML 编辑器，则仅插入文本数据，即使数据为 HTML。</span><span class="sxs-lookup"><span data-stu-id="67946-p187">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="67946-p188">如果 `html` 和该字段支持 HTML（主题不支持），则在 Outlook Web App 中应用当前样式，而在 Outlook 中应用默认样式。如果该字段是文本字段，则返回 `InvalidDataFormat` 错误。</span><span class="sxs-lookup"><span data-stu-id="67946-p188">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="67946-1391">如果未设置 `coercionType`，则结果取决于该字段：如果该字段是 HTML，则使用 HTML；如果该字段是文本，则使用纯文本。</span><span class="sxs-lookup"><span data-stu-id="67946-1391">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="67946-1392">function</span><span class="sxs-lookup"><span data-stu-id="67946-1392">function</span></span>||<span data-ttu-id="67946-1393">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="67946-1393">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="67946-1394">Requirements</span><span class="sxs-lookup"><span data-stu-id="67946-1394">Requirements</span></span>

|<span data-ttu-id="67946-1395">要求</span><span class="sxs-lookup"><span data-stu-id="67946-1395">Requirement</span></span>|<span data-ttu-id="67946-1396">值</span><span class="sxs-lookup"><span data-stu-id="67946-1396">Value</span></span>|
|---|---|
|[<span data-ttu-id="67946-1397">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="67946-1397">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="67946-1398">1.2</span><span class="sxs-lookup"><span data-stu-id="67946-1398">1.2</span></span>|
|[<span data-ttu-id="67946-1399">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="67946-1399">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="67946-1400">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="67946-1400">ReadWriteItem</span></span>|
|[<span data-ttu-id="67946-1401">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="67946-1401">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="67946-1402">撰写</span><span class="sxs-lookup"><span data-stu-id="67946-1402">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="67946-1403">示例</span><span class="sxs-lookup"><span data-stu-id="67946-1403">Example</span></span>

```
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```