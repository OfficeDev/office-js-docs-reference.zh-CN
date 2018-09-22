
# <a name="item"></a><span data-ttu-id="2a739-101">item</span><span class="sxs-lookup"><span data-stu-id="2a739-101">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="2a739-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="2a739-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="2a739-p101">`item` 命名空间用于访问当前选定的邮件、会议请求或约会。可以通过使用 [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) 属性确定 `item` 的类型。</span><span class="sxs-lookup"><span data-stu-id="2a739-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="2a739-105">Requirements</span><span class="sxs-lookup"><span data-stu-id="2a739-105">Requirements</span></span>

|<span data-ttu-id="2a739-106">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-106">Requirement</span></span>|<span data-ttu-id="2a739-107">值</span><span class="sxs-lookup"><span data-stu-id="2a739-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="2a739-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2a739-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2a739-109">1.0</span><span class="sxs-lookup"><span data-stu-id="2a739-109">1.0</span></span>|
|[<span data-ttu-id="2a739-110">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2a739-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2a739-111">受限</span><span class="sxs-lookup"><span data-stu-id="2a739-111">Restricted</span></span>|
|[<span data-ttu-id="2a739-112">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2a739-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2a739-113">Compose 或 Read</span><span class="sxs-lookup"><span data-stu-id="2a739-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="2a739-114">成员和方法</span><span class="sxs-lookup"><span data-stu-id="2a739-114">Members and methods</span></span>

| <span data-ttu-id="2a739-115">成员</span><span class="sxs-lookup"><span data-stu-id="2a739-115">Member</span></span> | <span data-ttu-id="2a739-116">类型</span><span class="sxs-lookup"><span data-stu-id="2a739-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="2a739-117">attachments</span><span class="sxs-lookup"><span data-stu-id="2a739-117">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails) | <span data-ttu-id="2a739-118">成员</span><span class="sxs-lookup"><span data-stu-id="2a739-118">Member</span></span> |
| [<span data-ttu-id="2a739-119">bcc</span><span class="sxs-lookup"><span data-stu-id="2a739-119">bcc</span></span>](#bcc-recipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="2a739-120">成员</span><span class="sxs-lookup"><span data-stu-id="2a739-120">Member</span></span> |
| [<span data-ttu-id="2a739-121">body</span><span class="sxs-lookup"><span data-stu-id="2a739-121">body</span></span>](#body-bodyjavascriptapioutlookofficebody) | <span data-ttu-id="2a739-122">成员</span><span class="sxs-lookup"><span data-stu-id="2a739-122">Member</span></span> |
| [<span data-ttu-id="2a739-123">cc</span><span class="sxs-lookup"><span data-stu-id="2a739-123">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="2a739-124">成员</span><span class="sxs-lookup"><span data-stu-id="2a739-124">Member</span></span> |
| [<span data-ttu-id="2a739-125">conversationId</span><span class="sxs-lookup"><span data-stu-id="2a739-125">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="2a739-126">成员</span><span class="sxs-lookup"><span data-stu-id="2a739-126">Member</span></span> |
| [<span data-ttu-id="2a739-127">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="2a739-127">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="2a739-128">成员</span><span class="sxs-lookup"><span data-stu-id="2a739-128">Member</span></span> |
| [<span data-ttu-id="2a739-129">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="2a739-129">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="2a739-130">成员</span><span class="sxs-lookup"><span data-stu-id="2a739-130">Member</span></span> |
| [<span data-ttu-id="2a739-131">end</span><span class="sxs-lookup"><span data-stu-id="2a739-131">end</span></span>](#end-datetimejavascriptapioutlookofficetime) | <span data-ttu-id="2a739-132">成员</span><span class="sxs-lookup"><span data-stu-id="2a739-132">Member</span></span> |
| [<span data-ttu-id="2a739-133">from</span><span class="sxs-lookup"><span data-stu-id="2a739-133">from</span></span>](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) | <span data-ttu-id="2a739-134">成员</span><span class="sxs-lookup"><span data-stu-id="2a739-134">Member</span></span> |
| [<span data-ttu-id="2a739-135">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="2a739-135">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="2a739-136">成员</span><span class="sxs-lookup"><span data-stu-id="2a739-136">Member</span></span> |
| [<span data-ttu-id="2a739-137">itemClass</span><span class="sxs-lookup"><span data-stu-id="2a739-137">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="2a739-138">成员</span><span class="sxs-lookup"><span data-stu-id="2a739-138">Member</span></span> |
| [<span data-ttu-id="2a739-139">itemId</span><span class="sxs-lookup"><span data-stu-id="2a739-139">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="2a739-140">成员</span><span class="sxs-lookup"><span data-stu-id="2a739-140">Member</span></span> |
| [<span data-ttu-id="2a739-141">itemType</span><span class="sxs-lookup"><span data-stu-id="2a739-141">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) | <span data-ttu-id="2a739-142">成员</span><span class="sxs-lookup"><span data-stu-id="2a739-142">Member</span></span> |
| [<span data-ttu-id="2a739-143">location</span><span class="sxs-lookup"><span data-stu-id="2a739-143">location</span></span>](#location-stringlocationjavascriptapioutlookofficelocation) | <span data-ttu-id="2a739-144">成员</span><span class="sxs-lookup"><span data-stu-id="2a739-144">Member</span></span> |
| [<span data-ttu-id="2a739-145">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="2a739-145">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="2a739-146">成员</span><span class="sxs-lookup"><span data-stu-id="2a739-146">Member</span></span> |
| [<span data-ttu-id="2a739-147">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="2a739-147">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages) | <span data-ttu-id="2a739-148">成员</span><span class="sxs-lookup"><span data-stu-id="2a739-148">Member</span></span> |
| [<span data-ttu-id="2a739-149">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="2a739-149">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="2a739-150">成员</span><span class="sxs-lookup"><span data-stu-id="2a739-150">Member</span></span> |
| [<span data-ttu-id="2a739-151">organizer</span><span class="sxs-lookup"><span data-stu-id="2a739-151">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer) | <span data-ttu-id="2a739-152">成员</span><span class="sxs-lookup"><span data-stu-id="2a739-152">Member</span></span> |
| [<span data-ttu-id="2a739-153">recurrence</span><span class="sxs-lookup"><span data-stu-id="2a739-153">recurrence</span></span>](#nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence) | <span data-ttu-id="2a739-154">成员</span><span class="sxs-lookup"><span data-stu-id="2a739-154">Member</span></span> |
| [<span data-ttu-id="2a739-155">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="2a739-155">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="2a739-156">成员</span><span class="sxs-lookup"><span data-stu-id="2a739-156">Member</span></span> |
| [<span data-ttu-id="2a739-157">sender</span><span class="sxs-lookup"><span data-stu-id="2a739-157">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) | <span data-ttu-id="2a739-158">成员</span><span class="sxs-lookup"><span data-stu-id="2a739-158">Member</span></span> |
| [<span data-ttu-id="2a739-159">可能指向</span><span class="sxs-lookup"><span data-stu-id="2a739-159">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="2a739-160">成员</span><span class="sxs-lookup"><span data-stu-id="2a739-160">Member</span></span> |
| [<span data-ttu-id="2a739-161">start</span><span class="sxs-lookup"><span data-stu-id="2a739-161">start</span></span>](#start-datetimejavascriptapioutlookofficetime) | <span data-ttu-id="2a739-162">成员</span><span class="sxs-lookup"><span data-stu-id="2a739-162">Member</span></span> |
| [<span data-ttu-id="2a739-163">subject</span><span class="sxs-lookup"><span data-stu-id="2a739-163">subject</span></span>](#subject-stringsubjectjavascriptapioutlookofficesubject) | <span data-ttu-id="2a739-164">成员</span><span class="sxs-lookup"><span data-stu-id="2a739-164">Member</span></span> |
| [<span data-ttu-id="2a739-165">to</span><span class="sxs-lookup"><span data-stu-id="2a739-165">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="2a739-166">成员</span><span class="sxs-lookup"><span data-stu-id="2a739-166">Member</span></span> |
| [<span data-ttu-id="2a739-167">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="2a739-167">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="2a739-168">方法</span><span class="sxs-lookup"><span data-stu-id="2a739-168">Method</span></span> |
| [<span data-ttu-id="2a739-169">addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="2a739-169">addFileAttachmentFromBase64Async</span></span>](#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) | <span data-ttu-id="2a739-170">方法</span><span class="sxs-lookup"><span data-stu-id="2a739-170">Method</span></span> |
| [<span data-ttu-id="2a739-171">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="2a739-171">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="2a739-172">方法</span><span class="sxs-lookup"><span data-stu-id="2a739-172">Method</span></span> |
| [<span data-ttu-id="2a739-173">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="2a739-173">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="2a739-174">方法</span><span class="sxs-lookup"><span data-stu-id="2a739-174">Method</span></span> |
| [<span data-ttu-id="2a739-175">close</span><span class="sxs-lookup"><span data-stu-id="2a739-175">close</span></span>](#close) | <span data-ttu-id="2a739-176">方法</span><span class="sxs-lookup"><span data-stu-id="2a739-176">Method</span></span> |
| [<span data-ttu-id="2a739-177">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="2a739-177">displayReplyAllForm</span></span>](#displayreplyallformformdata) | <span data-ttu-id="2a739-178">方法</span><span class="sxs-lookup"><span data-stu-id="2a739-178">Method</span></span> |
| [<span data-ttu-id="2a739-179">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="2a739-179">displayReplyForm</span></span>](#displayreplyformformdata) | <span data-ttu-id="2a739-180">方法</span><span class="sxs-lookup"><span data-stu-id="2a739-180">Method</span></span> |
| [<span data-ttu-id="2a739-181">getEntities</span><span class="sxs-lookup"><span data-stu-id="2a739-181">getEntities</span></span>](#getentities--entitiesjavascriptapioutlookofficeentities) | <span data-ttu-id="2a739-182">方法</span><span class="sxs-lookup"><span data-stu-id="2a739-182">Method</span></span> |
| [<span data-ttu-id="2a739-183">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="2a739-183">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion) | <span data-ttu-id="2a739-184">方法</span><span class="sxs-lookup"><span data-stu-id="2a739-184">Method</span></span> |
| [<span data-ttu-id="2a739-185">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="2a739-185">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion) | <span data-ttu-id="2a739-186">方法</span><span class="sxs-lookup"><span data-stu-id="2a739-186">Method</span></span> |
| [<span data-ttu-id="2a739-187">getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="2a739-187">getInitializationContextAsync</span></span>](#getinitializationcontextasyncoptions-callback) | <span data-ttu-id="2a739-188">方法</span><span class="sxs-lookup"><span data-stu-id="2a739-188">Method</span></span> |
| [<span data-ttu-id="2a739-189">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="2a739-189">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="2a739-190">方法</span><span class="sxs-lookup"><span data-stu-id="2a739-190">Method</span></span> |
| [<span data-ttu-id="2a739-191">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="2a739-191">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="2a739-192">方法</span><span class="sxs-lookup"><span data-stu-id="2a739-192">Method</span></span> |
| [<span data-ttu-id="2a739-193">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="2a739-193">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="2a739-194">方法</span><span class="sxs-lookup"><span data-stu-id="2a739-194">Method</span></span> |
| [<span data-ttu-id="2a739-195">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="2a739-195">getSelectedEntities</span></span>](#getselectedentities--entitiesjavascriptapioutlookofficeentities) | <span data-ttu-id="2a739-196">方法</span><span class="sxs-lookup"><span data-stu-id="2a739-196">Method</span></span> |
| [<span data-ttu-id="2a739-197">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="2a739-197">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="2a739-198">方法</span><span class="sxs-lookup"><span data-stu-id="2a739-198">Method</span></span> |
| [<span data-ttu-id="2a739-199">getSharedPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="2a739-199">getSharedPropertiesAsync</span></span>](#getsharedpropertiesasyncoptions-callback) | <span data-ttu-id="2a739-200">方法</span><span class="sxs-lookup"><span data-stu-id="2a739-200">Method</span></span> |
| [<span data-ttu-id="2a739-201">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="2a739-201">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="2a739-202">方法</span><span class="sxs-lookup"><span data-stu-id="2a739-202">Method</span></span> |
| [<span data-ttu-id="2a739-203">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="2a739-203">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="2a739-204">方法</span><span class="sxs-lookup"><span data-stu-id="2a739-204">Method</span></span> |
| [<span data-ttu-id="2a739-205">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="2a739-205">removeHandlerAsync</span></span>](#removehandlerasynceventtype-handler-options-callback) | <span data-ttu-id="2a739-206">方法</span><span class="sxs-lookup"><span data-stu-id="2a739-206">Method</span></span> |
| [<span data-ttu-id="2a739-207">saveAsync</span><span class="sxs-lookup"><span data-stu-id="2a739-207">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="2a739-208">方法</span><span class="sxs-lookup"><span data-stu-id="2a739-208">Method</span></span> |
| [<span data-ttu-id="2a739-209">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="2a739-209">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="2a739-210">方法</span><span class="sxs-lookup"><span data-stu-id="2a739-210">Method</span></span> |

### <a name="example"></a><span data-ttu-id="2a739-211">示例</span><span class="sxs-lookup"><span data-stu-id="2a739-211">Example</span></span>

<span data-ttu-id="2a739-212">以下 JavaScript 代码示例显示了如何访问 Outlook 中当前项目的 `subject` 属性。</span><span class="sxs-lookup"><span data-stu-id="2a739-212">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="2a739-213">成员</span><span class="sxs-lookup"><span data-stu-id="2a739-213">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="2a739-214">attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="2a739-214">attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="2a739-p102">获取项目的附件数组。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="2a739-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="2a739-217">某些类型的文件被阻止 Outlook 由于潜在的安全问题，并因此不返回。</span><span class="sxs-lookup"><span data-stu-id="2a739-217">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="2a739-218">有关详细信息，请参阅[在 Outlook 中被阻止的附件](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)。</span><span class="sxs-lookup"><span data-stu-id="2a739-218">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="2a739-219">类型:</span><span class="sxs-lookup"><span data-stu-id="2a739-219">Type:</span></span>

*   <span data-ttu-id="2a739-220">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="2a739-220">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="2a739-221">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-221">Requirements</span></span>

|<span data-ttu-id="2a739-222">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-222">Requirement</span></span>|<span data-ttu-id="2a739-223">值</span><span class="sxs-lookup"><span data-stu-id="2a739-223">Value</span></span>|
|---|---|
|[<span data-ttu-id="2a739-224">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2a739-224">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2a739-225">1.0</span><span class="sxs-lookup"><span data-stu-id="2a739-225">1.0</span></span>|
|[<span data-ttu-id="2a739-226">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2a739-226">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2a739-227">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2a739-227">ReadItem</span></span>|
|[<span data-ttu-id="2a739-228">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2a739-228">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2a739-229">阅读</span><span class="sxs-lookup"><span data-stu-id="2a739-229">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2a739-230">示例</span><span class="sxs-lookup"><span data-stu-id="2a739-230">Example</span></span>

<span data-ttu-id="2a739-231">以下代码使用当前项目上所有附件的详细信息构成 HTML 字符串。</span><span class="sxs-lookup"><span data-stu-id="2a739-231">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="2a739-232">bcc :[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="2a739-232">bcc :[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="2a739-233">获取一个对象，提供用于获取或更新邮件的密件抄送 （密件抄送副本） 行上的收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="2a739-233">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="2a739-234">仅限撰写模式。</span><span class="sxs-lookup"><span data-stu-id="2a739-234">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="2a739-235">类型：</span><span class="sxs-lookup"><span data-stu-id="2a739-235">Type:</span></span>

*   [<span data-ttu-id="2a739-236">收件人</span><span class="sxs-lookup"><span data-stu-id="2a739-236">Recipients</span></span>](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="2a739-237">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-237">Requirements</span></span>

|<span data-ttu-id="2a739-238">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-238">Requirement</span></span>|<span data-ttu-id="2a739-239">值</span><span class="sxs-lookup"><span data-stu-id="2a739-239">Value</span></span>|
|---|---|
|[<span data-ttu-id="2a739-240">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2a739-240">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2a739-241">1.1</span><span class="sxs-lookup"><span data-stu-id="2a739-241">1.1</span></span>|
|[<span data-ttu-id="2a739-242">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2a739-242">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2a739-243">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2a739-243">ReadItem</span></span>|
|[<span data-ttu-id="2a739-244">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2a739-244">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2a739-245">撰写</span><span class="sxs-lookup"><span data-stu-id="2a739-245">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="2a739-246">示例</span><span class="sxs-lookup"><span data-stu-id="2a739-246">Example</span></span>

```
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlookofficebody"></a><span data-ttu-id="2a739-247">body :[Body](/javascript/api/outlook/office.body)</span><span class="sxs-lookup"><span data-stu-id="2a739-247">body :[Body](/javascript/api/outlook/office.body)</span></span>

<span data-ttu-id="2a739-248">获取一个提供用于处理项目正文的方法的对象。</span><span class="sxs-lookup"><span data-stu-id="2a739-248">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="2a739-249">类型：</span><span class="sxs-lookup"><span data-stu-id="2a739-249">Type:</span></span>

*   [<span data-ttu-id="2a739-250">Body</span><span class="sxs-lookup"><span data-stu-id="2a739-250">Body</span></span>](/javascript/api/outlook/office.body)

##### <a name="requirements"></a><span data-ttu-id="2a739-251">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-251">Requirements</span></span>

|<span data-ttu-id="2a739-252">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-252">Requirement</span></span>|<span data-ttu-id="2a739-253">值</span><span class="sxs-lookup"><span data-stu-id="2a739-253">Value</span></span>|
|---|---|
|[<span data-ttu-id="2a739-254">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2a739-254">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2a739-255">1.1</span><span class="sxs-lookup"><span data-stu-id="2a739-255">1.1</span></span>|
|[<span data-ttu-id="2a739-256">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2a739-256">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2a739-257">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2a739-257">ReadItem</span></span>|
|[<span data-ttu-id="2a739-258">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2a739-258">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2a739-259">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="2a739-259">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="2a739-260">抄送： 数组。 <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[收件人](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="2a739-260">cc :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="2a739-261">提供对邮件的抄送 (cc) 收件人访问。</span><span class="sxs-lookup"><span data-stu-id="2a739-261">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="2a739-262">对象的类型和访问级别取决于当前项的模式。</span><span class="sxs-lookup"><span data-stu-id="2a739-262">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="2a739-263">阅读模式</span><span class="sxs-lookup"><span data-stu-id="2a739-263">Read mode</span></span>

<span data-ttu-id="2a739-p106">`cc` 属性返回包含邮件的**抄送**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="2a739-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="2a739-266">撰写模式</span><span class="sxs-lookup"><span data-stu-id="2a739-266">Compose mode</span></span>

<span data-ttu-id="2a739-267">`cc`属性返回`Recipients`提供方法用于获取或更新的邮件**Cc**行上的收件人对象。</span><span class="sxs-lookup"><span data-stu-id="2a739-267">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="2a739-268">类型:</span><span class="sxs-lookup"><span data-stu-id="2a739-268">Type:</span></span>

*   <span data-ttu-id="2a739-269">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="2a739-269">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="2a739-270">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-270">Requirements</span></span>

|<span data-ttu-id="2a739-271">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-271">Requirement</span></span>|<span data-ttu-id="2a739-272">值</span><span class="sxs-lookup"><span data-stu-id="2a739-272">Value</span></span>|
|---|---|
|[<span data-ttu-id="2a739-273">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2a739-273">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2a739-274">1.0</span><span class="sxs-lookup"><span data-stu-id="2a739-274">1.0</span></span>|
|[<span data-ttu-id="2a739-275">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2a739-275">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2a739-276">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2a739-276">ReadItem</span></span>|
|[<span data-ttu-id="2a739-277">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2a739-277">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2a739-278">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="2a739-278">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="2a739-279">示例</span><span class="sxs-lookup"><span data-stu-id="2a739-279">Example</span></span>

```
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="2a739-280">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="2a739-280">(nullable) conversationId :String</span></span>

<span data-ttu-id="2a739-281">获取包含特定消息的电子邮件会话的标识符。</span><span class="sxs-lookup"><span data-stu-id="2a739-281">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="2a739-p107">如果在阅读窗体或撰写窗体的回复中激活邮件应用程序，则此属性可以获得一个整数值。如果用户随后更改了回复邮件的主题（若发送回复），则该邮件的对话 ID 将改变且之前获取的值将不适用。</span><span class="sxs-lookup"><span data-stu-id="2a739-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="2a739-p108">对于撰写窗体的新项目，此属性获得一个 null 值。如果用户设置一个主题并保存该项目，`conversationId` 属性将返回一个值。</span><span class="sxs-lookup"><span data-stu-id="2a739-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="2a739-286">类型:</span><span class="sxs-lookup"><span data-stu-id="2a739-286">Type:</span></span>

*   <span data-ttu-id="2a739-287">String</span><span class="sxs-lookup"><span data-stu-id="2a739-287">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2a739-288">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-288">Requirements</span></span>

|<span data-ttu-id="2a739-289">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-289">Requirement</span></span>|<span data-ttu-id="2a739-290">值</span><span class="sxs-lookup"><span data-stu-id="2a739-290">Value</span></span>|
|---|---|
|[<span data-ttu-id="2a739-291">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2a739-291">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2a739-292">1.0</span><span class="sxs-lookup"><span data-stu-id="2a739-292">1.0</span></span>|
|[<span data-ttu-id="2a739-293">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2a739-293">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2a739-294">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2a739-294">ReadItem</span></span>|
|[<span data-ttu-id="2a739-295">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2a739-295">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2a739-296">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="2a739-296">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="2a739-297">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="2a739-297">dateTimeCreated :Date</span></span>

<span data-ttu-id="2a739-p109">获取项目创建的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="2a739-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="2a739-300">类型：</span><span class="sxs-lookup"><span data-stu-id="2a739-300">Type:</span></span>

*   <span data-ttu-id="2a739-301">日期</span><span class="sxs-lookup"><span data-stu-id="2a739-301">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="2a739-302">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-302">Requirements</span></span>

|<span data-ttu-id="2a739-303">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-303">Requirement</span></span>|<span data-ttu-id="2a739-304">值</span><span class="sxs-lookup"><span data-stu-id="2a739-304">Value</span></span>|
|---|---|
|[<span data-ttu-id="2a739-305">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2a739-305">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2a739-306">1.0</span><span class="sxs-lookup"><span data-stu-id="2a739-306">1.0</span></span>|
|[<span data-ttu-id="2a739-307">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2a739-307">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2a739-308">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2a739-308">ReadItem</span></span>|
|[<span data-ttu-id="2a739-309">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2a739-309">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2a739-310">阅读</span><span class="sxs-lookup"><span data-stu-id="2a739-310">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2a739-311">示例</span><span class="sxs-lookup"><span data-stu-id="2a739-311">Example</span></span>

```
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="2a739-312">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="2a739-312">dateTimeModified :Date</span></span>

<span data-ttu-id="2a739-p110">获取项目最近一次修改的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="2a739-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="2a739-315">此成员不支持适用于 iOS 的 Outlook 或 Outlook for Android。</span><span class="sxs-lookup"><span data-stu-id="2a739-315">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="2a739-316">类型:</span><span class="sxs-lookup"><span data-stu-id="2a739-316">Type:</span></span>

*   <span data-ttu-id="2a739-317">日期</span><span class="sxs-lookup"><span data-stu-id="2a739-317">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="2a739-318">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-318">Requirements</span></span>

|<span data-ttu-id="2a739-319">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-319">Requirement</span></span>|<span data-ttu-id="2a739-320">值</span><span class="sxs-lookup"><span data-stu-id="2a739-320">Value</span></span>|
|---|---|
|[<span data-ttu-id="2a739-321">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2a739-321">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2a739-322">1.0</span><span class="sxs-lookup"><span data-stu-id="2a739-322">1.0</span></span>|
|[<span data-ttu-id="2a739-323">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2a739-323">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2a739-324">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2a739-324">ReadItem</span></span>|
|[<span data-ttu-id="2a739-325">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2a739-325">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2a739-326">阅读</span><span class="sxs-lookup"><span data-stu-id="2a739-326">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2a739-327">示例</span><span class="sxs-lookup"><span data-stu-id="2a739-327">Example</span></span>

```
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="2a739-328">end :Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="2a739-328">end :Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="2a739-329">获取或设置约会结束的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="2a739-329">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="2a739-p111">将 `end` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) 方法将 end 属性值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="2a739-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="2a739-332">阅读模式</span><span class="sxs-lookup"><span data-stu-id="2a739-332">Read mode</span></span>

<span data-ttu-id="2a739-333">`end` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="2a739-333">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="2a739-334">撰写模式</span><span class="sxs-lookup"><span data-stu-id="2a739-334">Compose mode</span></span>

<span data-ttu-id="2a739-335">`end` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="2a739-335">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="2a739-336">使用 [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) 方法设置结束时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="2a739-336">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="2a739-337">类型：</span><span class="sxs-lookup"><span data-stu-id="2a739-337">Type:</span></span>

*   <span data-ttu-id="2a739-338">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="2a739-338">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="2a739-339">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-339">Requirements</span></span>

|<span data-ttu-id="2a739-340">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-340">Requirement</span></span>|<span data-ttu-id="2a739-341">值</span><span class="sxs-lookup"><span data-stu-id="2a739-341">Value</span></span>|
|---|---|
|[<span data-ttu-id="2a739-342">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2a739-342">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2a739-343">1.0</span><span class="sxs-lookup"><span data-stu-id="2a739-343">1.0</span></span>|
|[<span data-ttu-id="2a739-344">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2a739-344">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2a739-345">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2a739-345">ReadItem</span></span>|
|[<span data-ttu-id="2a739-346">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2a739-346">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2a739-347">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="2a739-347">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="2a739-348">示例</span><span class="sxs-lookup"><span data-stu-id="2a739-348">Example</span></span>

<span data-ttu-id="2a739-349">以下示例通过使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) 方法，设置撰写模式下约会的结束时间。</span><span class="sxs-lookup"><span data-stu-id="2a739-349">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom"></a><span data-ttu-id="2a739-350">从：[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[从](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="2a739-350">from :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span></span>

<span data-ttu-id="2a739-351">获取邮件发件人的电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="2a739-351">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="2a739-p112">`from` 和 [`sender`](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="2a739-p112">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="2a739-354">`recipientType`属性`EmailAddressDetails`对象在`from`属性是`undefined`。</span><span class="sxs-lookup"><span data-stu-id="2a739-354">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="2a739-355">阅读模式</span><span class="sxs-lookup"><span data-stu-id="2a739-355">Read mode</span></span>

<span data-ttu-id="2a739-356">`from`属性返回`EmailAddressDetails`对象。</span><span class="sxs-lookup"><span data-stu-id="2a739-356">The `from` property returns an `EmailAddressDetails` object.</span></span>

```
var subject = Office.context.mailbox.item.from;
```

##### <a name="compose-mode"></a><span data-ttu-id="2a739-357">撰写模式</span><span class="sxs-lookup"><span data-stu-id="2a739-357">Compose mode</span></span>

<span data-ttu-id="2a739-358">`from`属性返回`From`对象，它提供一个方法来获取值。</span><span class="sxs-lookup"><span data-stu-id="2a739-358">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="2a739-359">类型:</span><span class="sxs-lookup"><span data-stu-id="2a739-359">Type:</span></span>

*   <span data-ttu-id="2a739-360">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [从](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="2a739-360">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="2a739-361">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-361">Requirements</span></span>

|<span data-ttu-id="2a739-362">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-362">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="2a739-363">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2a739-363">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2a739-364">1.0</span><span class="sxs-lookup"><span data-stu-id="2a739-364">1.0</span></span>|<span data-ttu-id="2a739-365">1.7</span><span class="sxs-lookup"><span data-stu-id="2a739-365">1.7</span></span>|
|[<span data-ttu-id="2a739-366">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2a739-366">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2a739-367">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2a739-367">ReadItem</span></span>|<span data-ttu-id="2a739-368">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="2a739-368">ReadWriteItem</span></span>|
|[<span data-ttu-id="2a739-369">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2a739-369">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2a739-370">Read</span><span class="sxs-lookup"><span data-stu-id="2a739-370">Read</span></span>|<span data-ttu-id="2a739-371">撰写</span><span class="sxs-lookup"><span data-stu-id="2a739-371">Compose</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="2a739-372">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="2a739-372">internetMessageId :String</span></span>

<span data-ttu-id="2a739-p113">获取电子邮件的 Internet 消息标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="2a739-p113">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="2a739-375">类型：</span><span class="sxs-lookup"><span data-stu-id="2a739-375">Type:</span></span>

*   <span data-ttu-id="2a739-376">String</span><span class="sxs-lookup"><span data-stu-id="2a739-376">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2a739-377">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-377">Requirements</span></span>

|<span data-ttu-id="2a739-378">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-378">Requirement</span></span>|<span data-ttu-id="2a739-379">值</span><span class="sxs-lookup"><span data-stu-id="2a739-379">Value</span></span>|
|---|---|
|[<span data-ttu-id="2a739-380">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2a739-380">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2a739-381">1.0</span><span class="sxs-lookup"><span data-stu-id="2a739-381">1.0</span></span>|
|[<span data-ttu-id="2a739-382">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2a739-382">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2a739-383">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2a739-383">ReadItem</span></span>|
|[<span data-ttu-id="2a739-384">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2a739-384">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2a739-385">阅读</span><span class="sxs-lookup"><span data-stu-id="2a739-385">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2a739-386">示例</span><span class="sxs-lookup"><span data-stu-id="2a739-386">Example</span></span>

```
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="2a739-387">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="2a739-387">itemClass :String</span></span>

<span data-ttu-id="2a739-p114">获取选定项目的 Exchange Web 服务项目类。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="2a739-p114">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="2a739-p115">`itemClass` 属性指定所选项目的邮件类别。以下是邮件或约会项目的默认邮件类别。</span><span class="sxs-lookup"><span data-stu-id="2a739-p115">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="2a739-392">类型</span><span class="sxs-lookup"><span data-stu-id="2a739-392">Type</span></span>|<span data-ttu-id="2a739-393">说明</span><span class="sxs-lookup"><span data-stu-id="2a739-393">Description</span></span>|<span data-ttu-id="2a739-394">项目类</span><span class="sxs-lookup"><span data-stu-id="2a739-394">item class</span></span>|
|---|---|---|
|<span data-ttu-id="2a739-395">约会项目</span><span class="sxs-lookup"><span data-stu-id="2a739-395">Appointment items</span></span>|<span data-ttu-id="2a739-396">这些是项目类 `IPM.Appointment` 或 `IPM.Appointment.Occurence` 的日历项目。</span><span class="sxs-lookup"><span data-stu-id="2a739-396">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurence`|
|<span data-ttu-id="2a739-397">邮件项目</span><span class="sxs-lookup"><span data-stu-id="2a739-397">Message items</span></span>|<span data-ttu-id="2a739-398">这些项目包括具有默认邮件类别 `IPM.Note` 的电子邮件，以及使用 `IPM.Schedule.Meeting` 作为基础邮件类别的会议请求、响应和取消。</span><span class="sxs-lookup"><span data-stu-id="2a739-398">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="2a739-399">你可以创建用于扩展默认邮件类别的自定义邮件类别，例如，自定义约会邮件类别 `IPM.Appointment.Contoso`。</span><span class="sxs-lookup"><span data-stu-id="2a739-399">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="2a739-400">类型：</span><span class="sxs-lookup"><span data-stu-id="2a739-400">Type:</span></span>

*   <span data-ttu-id="2a739-401">String</span><span class="sxs-lookup"><span data-stu-id="2a739-401">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2a739-402">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-402">Requirements</span></span>

|<span data-ttu-id="2a739-403">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-403">Requirement</span></span>|<span data-ttu-id="2a739-404">值</span><span class="sxs-lookup"><span data-stu-id="2a739-404">Value</span></span>|
|---|---|
|[<span data-ttu-id="2a739-405">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2a739-405">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2a739-406">1.0</span><span class="sxs-lookup"><span data-stu-id="2a739-406">1.0</span></span>|
|[<span data-ttu-id="2a739-407">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2a739-407">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2a739-408">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2a739-408">ReadItem</span></span>|
|[<span data-ttu-id="2a739-409">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2a739-409">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2a739-410">阅读</span><span class="sxs-lookup"><span data-stu-id="2a739-410">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2a739-411">示例</span><span class="sxs-lookup"><span data-stu-id="2a739-411">Example</span></span>

```
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="2a739-412">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="2a739-412">(nullable) itemId :String</span></span>

<span data-ttu-id="2a739-p116">获取当前项目的 Exchange Web 服务项目标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="2a739-p116">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="2a739-415">`itemId` 属性返回的标识符与 Exchange Web 服务项目标识符相同。</span><span class="sxs-lookup"><span data-stu-id="2a739-415">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="2a739-416">`itemId`属性不是与 Outlook 条目 ID 或使用 Outlook REST API 的 ID。</span><span class="sxs-lookup"><span data-stu-id="2a739-416">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="2a739-417">使用此值的 REST API 调用之前，应将其转换使用[Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string)。</span><span class="sxs-lookup"><span data-stu-id="2a739-417">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="2a739-418">有关详细信息，请参阅[使用 Outlook REST Api 从 Outlook 外接程序](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id)。</span><span class="sxs-lookup"><span data-stu-id="2a739-418">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="2a739-p118">`itemId` 属性在撰写模式下不可用。如果需要项目标识符，[`saveAsync`](#saveasyncoptions-callback) 方法可用于将项目保存到存储，这将在回调函数的 [`AsyncResult.value`](/javascript/api/office/office.asyncresult) 参数中返回项目标识符。</span><span class="sxs-lookup"><span data-stu-id="2a739-p118">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="2a739-421">类型:</span><span class="sxs-lookup"><span data-stu-id="2a739-421">Type:</span></span>

*   <span data-ttu-id="2a739-422">String</span><span class="sxs-lookup"><span data-stu-id="2a739-422">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2a739-423">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-423">Requirements</span></span>

|<span data-ttu-id="2a739-424">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-424">Requirement</span></span>|<span data-ttu-id="2a739-425">值</span><span class="sxs-lookup"><span data-stu-id="2a739-425">Value</span></span>|
|---|---|
|[<span data-ttu-id="2a739-426">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2a739-426">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2a739-427">1.0</span><span class="sxs-lookup"><span data-stu-id="2a739-427">1.0</span></span>|
|[<span data-ttu-id="2a739-428">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2a739-428">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2a739-429">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2a739-429">ReadItem</span></span>|
|[<span data-ttu-id="2a739-430">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2a739-430">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2a739-431">阅读</span><span class="sxs-lookup"><span data-stu-id="2a739-431">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2a739-432">示例</span><span class="sxs-lookup"><span data-stu-id="2a739-432">Example</span></span>

<span data-ttu-id="2a739-p119">以下代码检查项目标识符是否存在。如果 `itemId` 属性返回 `null` 或 `undefined`，则将项目保存到存储，并从异步结果中获取项目标识符。</span><span class="sxs-lookup"><span data-stu-id="2a739-p119">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype"></a><span data-ttu-id="2a739-435">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="2a739-435">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="2a739-436">获取实例表示的项的类型。</span><span class="sxs-lookup"><span data-stu-id="2a739-436">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="2a739-437">`itemType` 属性返回其中一个 `ItemType` 枚举值，指示 `item` 对象实例是邮件还是约会。</span><span class="sxs-lookup"><span data-stu-id="2a739-437">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="2a739-438">类型：</span><span class="sxs-lookup"><span data-stu-id="2a739-438">Type:</span></span>

*   [<span data-ttu-id="2a739-439">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="2a739-439">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="2a739-440">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-440">Requirements</span></span>

|<span data-ttu-id="2a739-441">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-441">Requirement</span></span>|<span data-ttu-id="2a739-442">值</span><span class="sxs-lookup"><span data-stu-id="2a739-442">Value</span></span>|
|---|---|
|[<span data-ttu-id="2a739-443">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2a739-443">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2a739-444">1.0</span><span class="sxs-lookup"><span data-stu-id="2a739-444">1.0</span></span>|
|[<span data-ttu-id="2a739-445">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2a739-445">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2a739-446">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2a739-446">ReadItem</span></span>|
|[<span data-ttu-id="2a739-447">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2a739-447">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2a739-448">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="2a739-448">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="2a739-449">示例</span><span class="sxs-lookup"><span data-stu-id="2a739-449">Example</span></span>

```
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlookofficelocation"></a><span data-ttu-id="2a739-450">location :String|[Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="2a739-450">location :String|[Location](/javascript/api/outlook/office.location)</span></span>

<span data-ttu-id="2a739-451">获取或设置约会的位置。</span><span class="sxs-lookup"><span data-stu-id="2a739-451">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="2a739-452">阅读模式</span><span class="sxs-lookup"><span data-stu-id="2a739-452">Read mode</span></span>

<span data-ttu-id="2a739-453">`location` 属性返回一个包含约会位置的字符串。</span><span class="sxs-lookup"><span data-stu-id="2a739-453">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="2a739-454">撰写模式</span><span class="sxs-lookup"><span data-stu-id="2a739-454">Compose mode</span></span>

<span data-ttu-id="2a739-455">`location` 属性返回一个 `Location` 对象，该对象提供用于获取和设置约会位置的方法。</span><span class="sxs-lookup"><span data-stu-id="2a739-455">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="2a739-456">类型：</span><span class="sxs-lookup"><span data-stu-id="2a739-456">Type:</span></span>

*   <span data-ttu-id="2a739-457">String | [Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="2a739-457">String | [Location](/javascript/api/outlook/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="2a739-458">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-458">Requirements</span></span>

|<span data-ttu-id="2a739-459">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-459">Requirement</span></span>|<span data-ttu-id="2a739-460">值</span><span class="sxs-lookup"><span data-stu-id="2a739-460">Value</span></span>|
|---|---|
|[<span data-ttu-id="2a739-461">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2a739-461">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2a739-462">1.0</span><span class="sxs-lookup"><span data-stu-id="2a739-462">1.0</span></span>|
|[<span data-ttu-id="2a739-463">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2a739-463">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2a739-464">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2a739-464">ReadItem</span></span>|
|[<span data-ttu-id="2a739-465">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2a739-465">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2a739-466">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="2a739-466">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="2a739-467">示例</span><span class="sxs-lookup"><span data-stu-id="2a739-467">Example</span></span>

```
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="2a739-468">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="2a739-468">normalizedSubject :String</span></span>

<span data-ttu-id="2a739-p120">获取删除了所有前缀（包括 `RE:` 和 `FWD:`）的项目主题。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="2a739-p120">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="2a739-p121">normalizedSubject 属性获取包含由电子邮件程序添加的任何标准前缀（如 `RE:` 和 `FW:`）的项目主题。若要获取包含完整前缀的项目主题，请使用 [`subject`](#subject-stringsubjectjavascriptapioutlookofficesubject) 属性。</span><span class="sxs-lookup"><span data-stu-id="2a739-p121">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlookofficesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="2a739-473">类型:</span><span class="sxs-lookup"><span data-stu-id="2a739-473">Type:</span></span>

*   <span data-ttu-id="2a739-474">String</span><span class="sxs-lookup"><span data-stu-id="2a739-474">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2a739-475">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-475">Requirements</span></span>

|<span data-ttu-id="2a739-476">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-476">Requirement</span></span>|<span data-ttu-id="2a739-477">值</span><span class="sxs-lookup"><span data-stu-id="2a739-477">Value</span></span>|
|---|---|
|[<span data-ttu-id="2a739-478">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2a739-478">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2a739-479">1.0</span><span class="sxs-lookup"><span data-stu-id="2a739-479">1.0</span></span>|
|[<span data-ttu-id="2a739-480">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2a739-480">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2a739-481">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2a739-481">ReadItem</span></span>|
|[<span data-ttu-id="2a739-482">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2a739-482">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2a739-483">阅读</span><span class="sxs-lookup"><span data-stu-id="2a739-483">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2a739-484">示例</span><span class="sxs-lookup"><span data-stu-id="2a739-484">Example</span></span>

```
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages"></a><span data-ttu-id="2a739-485">notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="2a739-485">notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span></span>

<span data-ttu-id="2a739-486">获取项目的通知邮件。</span><span class="sxs-lookup"><span data-stu-id="2a739-486">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="2a739-487">类型：</span><span class="sxs-lookup"><span data-stu-id="2a739-487">Type:</span></span>

*   [<span data-ttu-id="2a739-488">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="2a739-488">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="2a739-489">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-489">Requirements</span></span>

|<span data-ttu-id="2a739-490">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-490">Requirement</span></span>|<span data-ttu-id="2a739-491">值</span><span class="sxs-lookup"><span data-stu-id="2a739-491">Value</span></span>|
|---|---|
|[<span data-ttu-id="2a739-492">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2a739-492">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2a739-493">1.3</span><span class="sxs-lookup"><span data-stu-id="2a739-493">1.3</span></span>|
|[<span data-ttu-id="2a739-494">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2a739-494">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2a739-495">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2a739-495">ReadItem</span></span>|
|[<span data-ttu-id="2a739-496">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2a739-496">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2a739-497">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="2a739-497">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="2a739-498">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="2a739-498">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="2a739-499">提供对事件的可选与会者的访问。</span><span class="sxs-lookup"><span data-stu-id="2a739-499">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="2a739-500">对象的类型和访问级别取决于当前项的模式。</span><span class="sxs-lookup"><span data-stu-id="2a739-500">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="2a739-501">阅读模式</span><span class="sxs-lookup"><span data-stu-id="2a739-501">Read mode</span></span>

<span data-ttu-id="2a739-502">`optionalAttendees` 属性返回一个数组，其中包含每个可选与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="2a739-502">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="2a739-503">撰写模式</span><span class="sxs-lookup"><span data-stu-id="2a739-503">Compose mode</span></span>

<span data-ttu-id="2a739-504">`optionalAttendees`属性返回`Recipients`对象，它提供用于获取或更新会议的可选与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="2a739-504">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="2a739-505">类型：</span><span class="sxs-lookup"><span data-stu-id="2a739-505">Type:</span></span>

*   <span data-ttu-id="2a739-506">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="2a739-506">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="2a739-507">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-507">Requirements</span></span>

|<span data-ttu-id="2a739-508">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-508">Requirement</span></span>|<span data-ttu-id="2a739-509">值</span><span class="sxs-lookup"><span data-stu-id="2a739-509">Value</span></span>|
|---|---|
|[<span data-ttu-id="2a739-510">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2a739-510">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2a739-511">1.0</span><span class="sxs-lookup"><span data-stu-id="2a739-511">1.0</span></span>|
|[<span data-ttu-id="2a739-512">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2a739-512">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2a739-513">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2a739-513">ReadItem</span></span>|
|[<span data-ttu-id="2a739-514">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2a739-514">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2a739-515">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="2a739-515">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="2a739-516">示例</span><span class="sxs-lookup"><span data-stu-id="2a739-516">Example</span></span>

```
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer"></a><span data-ttu-id="2a739-517">组织者：[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[组织者](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="2a739-517">organizer :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span></span>

<span data-ttu-id="2a739-518">获取指定会议组织者的电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="2a739-518">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="2a739-519">阅读模式</span><span class="sxs-lookup"><span data-stu-id="2a739-519">Read mode</span></span>

<span data-ttu-id="2a739-520">`organizer`属性返回一个代表会议组织者的[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)对象。</span><span class="sxs-lookup"><span data-stu-id="2a739-520">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="2a739-521">撰写模式</span><span class="sxs-lookup"><span data-stu-id="2a739-521">Compose mode</span></span>

<span data-ttu-id="2a739-522">`organizer`属性返回一个[管理器](/javascript/api/outlook/office.organizer)对象，提供要获取组织者值的方法。</span><span class="sxs-lookup"><span data-stu-id="2a739-522">The `organizer` property returns an [Organizer](/javascript/api/outlook/office.organizer) object that provides a method to get the organizer value.</span></span>

##### <a name="type"></a><span data-ttu-id="2a739-523">类型:</span><span class="sxs-lookup"><span data-stu-id="2a739-523">Type:</span></span>

*   <span data-ttu-id="2a739-524">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [组织者](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="2a739-524">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="2a739-525">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-525">Requirements</span></span>

|<span data-ttu-id="2a739-526">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-526">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="2a739-527">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2a739-527">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2a739-528">1.0</span><span class="sxs-lookup"><span data-stu-id="2a739-528">1.0</span></span>|<span data-ttu-id="2a739-529">1.7</span><span class="sxs-lookup"><span data-stu-id="2a739-529">1.7</span></span>|
|[<span data-ttu-id="2a739-530">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2a739-530">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2a739-531">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2a739-531">ReadItem</span></span>|<span data-ttu-id="2a739-532">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="2a739-532">ReadWriteItem</span></span>|
|[<span data-ttu-id="2a739-533">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2a739-533">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2a739-534">Read</span><span class="sxs-lookup"><span data-stu-id="2a739-534">Read</span></span>|<span data-ttu-id="2a739-535">撰写</span><span class="sxs-lookup"><span data-stu-id="2a739-535">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="2a739-536">示例</span><span class="sxs-lookup"><span data-stu-id="2a739-536">Example</span></span>

```
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence"></a><span data-ttu-id="2a739-537">(可以为 null) 的定期：[定期](/javascript/api/outlook/office.recurrence)</span><span class="sxs-lookup"><span data-stu-id="2a739-537">(nullable) recurrence :[Recurrence](/javascript/api/outlook/office.recurrence)</span></span>

<span data-ttu-id="2a739-538">获取或设置约会的定期模式。</span><span class="sxs-lookup"><span data-stu-id="2a739-538">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="2a739-539">获取会议请求的定期模式。</span><span class="sxs-lookup"><span data-stu-id="2a739-539">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="2a739-540">阅读和撰写的约会项目的模式。</span><span class="sxs-lookup"><span data-stu-id="2a739-540">Read and compose modes for appointment items.</span></span> <span data-ttu-id="2a739-541">读取模式的会议请求项目。</span><span class="sxs-lookup"><span data-stu-id="2a739-541">Read mode for meeting request items.</span></span>

<span data-ttu-id="2a739-542">`recurrence`属性返回的定期约会或会议请求的[定期](/javascript/api/outlook/office.recurrence)对象，如果项是一系列或一系列中的实例。</span><span class="sxs-lookup"><span data-stu-id="2a739-542">The `recurrence` property returns a [recurrence](/javascript/api/outlook/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="2a739-543">`null`将返回单个约会和会议请求的单个约会。</span><span class="sxs-lookup"><span data-stu-id="2a739-543">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="2a739-544">`undefined`将返回不是会议请求的消息。</span><span class="sxs-lookup"><span data-stu-id="2a739-544">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="2a739-545">注意： 会议请求具有`itemClass`IPM 的值。Schedule.Meeting.Request。</span><span class="sxs-lookup"><span data-stu-id="2a739-545">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="2a739-546">注意： 如果定期对象是`null`，这表明该对象是一个约会或会议请求的一个约会，不属于一系列。</span><span class="sxs-lookup"><span data-stu-id="2a739-546">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="type"></a><span data-ttu-id="2a739-547">类型:</span><span class="sxs-lookup"><span data-stu-id="2a739-547">Type:</span></span>

* [<span data-ttu-id="2a739-548">定期</span><span class="sxs-lookup"><span data-stu-id="2a739-548">Recurrence</span></span>](/javascript/api/outlook/office.recurrence)

|<span data-ttu-id="2a739-549">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-549">Requirement</span></span>|<span data-ttu-id="2a739-550">值</span><span class="sxs-lookup"><span data-stu-id="2a739-550">Value</span></span>|
|---|---|
|[<span data-ttu-id="2a739-551">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2a739-551">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2a739-552">1.7</span><span class="sxs-lookup"><span data-stu-id="2a739-552">1.7</span></span>|
|[<span data-ttu-id="2a739-553">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2a739-553">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2a739-554">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2a739-554">ReadItem</span></span>|
|[<span data-ttu-id="2a739-555">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2a739-555">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2a739-556">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="2a739-556">Compose or read</span></span>|

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="2a739-557">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="2a739-557">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="2a739-558">提供对事件的必需与会者的访问。</span><span class="sxs-lookup"><span data-stu-id="2a739-558">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="2a739-559">对象的类型和访问级别取决于当前项的模式。</span><span class="sxs-lookup"><span data-stu-id="2a739-559">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="2a739-560">阅读模式</span><span class="sxs-lookup"><span data-stu-id="2a739-560">Read mode</span></span>

<span data-ttu-id="2a739-561">`requiredAttendees` 属性返回一个数组，其中包含每个必需与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="2a739-561">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="2a739-562">撰写模式</span><span class="sxs-lookup"><span data-stu-id="2a739-562">Compose mode</span></span>

<span data-ttu-id="2a739-563">`requiredAttendees`属性返回`Recipients`对象，它提供用于获取或更新会议的必需的与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="2a739-563">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="2a739-564">类型：</span><span class="sxs-lookup"><span data-stu-id="2a739-564">Type:</span></span>

*   <span data-ttu-id="2a739-565">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="2a739-565">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="2a739-566">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-566">Requirements</span></span>

|<span data-ttu-id="2a739-567">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-567">Requirement</span></span>|<span data-ttu-id="2a739-568">值</span><span class="sxs-lookup"><span data-stu-id="2a739-568">Value</span></span>|
|---|---|
|[<span data-ttu-id="2a739-569">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2a739-569">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2a739-570">1.0</span><span class="sxs-lookup"><span data-stu-id="2a739-570">1.0</span></span>|
|[<span data-ttu-id="2a739-571">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2a739-571">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2a739-572">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2a739-572">ReadItem</span></span>|
|[<span data-ttu-id="2a739-573">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2a739-573">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2a739-574">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="2a739-574">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="2a739-575">示例</span><span class="sxs-lookup"><span data-stu-id="2a739-575">Example</span></span>

```
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails"></a><span data-ttu-id="2a739-576">sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="2a739-576">sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span></span>

<span data-ttu-id="2a739-p126">获取电子邮件发件人的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="2a739-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="2a739-p127">[`from`](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) 和 `sender` 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="2a739-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="2a739-581">`recipientType`属性`EmailAddressDetails`对象在`sender`属性是`undefined`。</span><span class="sxs-lookup"><span data-stu-id="2a739-581">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="2a739-582">类型:</span><span class="sxs-lookup"><span data-stu-id="2a739-582">Type:</span></span>

*   [<span data-ttu-id="2a739-583">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="2a739-583">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="2a739-584">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-584">Requirements</span></span>

|<span data-ttu-id="2a739-585">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-585">Requirement</span></span>|<span data-ttu-id="2a739-586">值</span><span class="sxs-lookup"><span data-stu-id="2a739-586">Value</span></span>|
|---|---|
|[<span data-ttu-id="2a739-587">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2a739-587">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2a739-588">1.0</span><span class="sxs-lookup"><span data-stu-id="2a739-588">1.0</span></span>|
|[<span data-ttu-id="2a739-589">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2a739-589">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2a739-590">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2a739-590">ReadItem</span></span>|
|[<span data-ttu-id="2a739-591">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2a739-591">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2a739-592">阅读</span><span class="sxs-lookup"><span data-stu-id="2a739-592">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2a739-593">示例</span><span class="sxs-lookup"><span data-stu-id="2a739-593">Example</span></span>

```
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="2a739-594">(nullable) 可能指向： 字符串</span><span class="sxs-lookup"><span data-stu-id="2a739-594">(nullable) seriesId :String</span></span>

<span data-ttu-id="2a739-595">获取实例所属的系列的 id。</span><span class="sxs-lookup"><span data-stu-id="2a739-595">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="2a739-596">在 OWA 和 Outlook 中，`seriesId`返回父 （系列） 项此项目所属的 Exchange Web Services (EWS) ID。</span><span class="sxs-lookup"><span data-stu-id="2a739-596">In OWA and Outlook, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="2a739-597">但是，在 iOS 和 Android，`seriesId`返回父项的 REST ID。</span><span class="sxs-lookup"><span data-stu-id="2a739-597">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="2a739-598">`seriesId` 属性返回的标识符与 Exchange Web 服务项目标识符相同。</span><span class="sxs-lookup"><span data-stu-id="2a739-598">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="2a739-599">`seriesId`属性不是与 Outlook Id 使用 Outlook REST API 相同。</span><span class="sxs-lookup"><span data-stu-id="2a739-599">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="2a739-600">使用此值的 REST API 调用之前，应将其转换使用[Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string)。</span><span class="sxs-lookup"><span data-stu-id="2a739-600">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="2a739-601">有关详细信息，请参阅[使用 Outlook REST Api 从 Outlook 外接程序](https://docs.microsoft.com/outlook/add-ins/use-rest-api)。</span><span class="sxs-lookup"><span data-stu-id="2a739-601">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="2a739-602">`seriesId`属性返回`null`如不具有父项的项单个系列项目、 约会或会议请求，并返回`undefined`不会议请求的其他任何项。</span><span class="sxs-lookup"><span data-stu-id="2a739-602">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="2a739-603">类型:</span><span class="sxs-lookup"><span data-stu-id="2a739-603">Type:</span></span>

* <span data-ttu-id="2a739-604">String</span><span class="sxs-lookup"><span data-stu-id="2a739-604">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2a739-605">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-605">Requirements</span></span>

|<span data-ttu-id="2a739-606">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-606">Requirement</span></span>|<span data-ttu-id="2a739-607">值</span><span class="sxs-lookup"><span data-stu-id="2a739-607">Value</span></span>|
|---|---|
|[<span data-ttu-id="2a739-608">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2a739-608">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2a739-609">1.7</span><span class="sxs-lookup"><span data-stu-id="2a739-609">1.7</span></span>|
|[<span data-ttu-id="2a739-610">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2a739-610">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2a739-611">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2a739-611">ReadItem</span></span>|
|[<span data-ttu-id="2a739-612">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2a739-612">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2a739-613">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="2a739-613">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="2a739-614">示例</span><span class="sxs-lookup"><span data-stu-id="2a739-614">Example</span></span>

```
var seriesId = Office.context.mailbox.item.seriesId;
var isSeries = (seriesId == null);
```

####  <a name="start-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="2a739-615">start :Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="2a739-615">start :Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="2a739-616">获取或设置约会开始的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="2a739-616">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="2a739-p130">将 `start` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) 方法将该值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="2a739-p130">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="2a739-619">阅读模式</span><span class="sxs-lookup"><span data-stu-id="2a739-619">Read mode</span></span>

<span data-ttu-id="2a739-620">`start` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="2a739-620">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="2a739-621">撰写模式</span><span class="sxs-lookup"><span data-stu-id="2a739-621">Compose mode</span></span>

<span data-ttu-id="2a739-622">`start` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="2a739-622">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="2a739-623">使用 [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) 方法设置开始时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="2a739-623">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="2a739-624">类型：</span><span class="sxs-lookup"><span data-stu-id="2a739-624">Type:</span></span>

*   <span data-ttu-id="2a739-625">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="2a739-625">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="2a739-626">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-626">Requirements</span></span>

|<span data-ttu-id="2a739-627">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-627">Requirement</span></span>|<span data-ttu-id="2a739-628">值</span><span class="sxs-lookup"><span data-stu-id="2a739-628">Value</span></span>|
|---|---|
|[<span data-ttu-id="2a739-629">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2a739-629">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2a739-630">1.0</span><span class="sxs-lookup"><span data-stu-id="2a739-630">1.0</span></span>|
|[<span data-ttu-id="2a739-631">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2a739-631">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2a739-632">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2a739-632">ReadItem</span></span>|
|[<span data-ttu-id="2a739-633">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2a739-633">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2a739-634">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="2a739-634">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="2a739-635">示例</span><span class="sxs-lookup"><span data-stu-id="2a739-635">Example</span></span>

<span data-ttu-id="2a739-636">以下示例通过使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) 方法，设置撰写模式下约会的开始时间。</span><span class="sxs-lookup"><span data-stu-id="2a739-636">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

####  <a name="subject-stringsubjectjavascriptapioutlookofficesubject"></a><span data-ttu-id="2a739-637">subject :String|[Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="2a739-637">subject :String|[Subject](/javascript/api/outlook/office.subject)</span></span>

<span data-ttu-id="2a739-638">获取或设置显示在项目的主题字段中的说明。</span><span class="sxs-lookup"><span data-stu-id="2a739-638">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="2a739-639">`subject` 属性获取或设置由电子邮件服务器发送项目时的整个主题。</span><span class="sxs-lookup"><span data-stu-id="2a739-639">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="2a739-640">阅读模式</span><span class="sxs-lookup"><span data-stu-id="2a739-640">Read mode</span></span>

<span data-ttu-id="2a739-p131">`subject` 属性返回一个字符串。使用 [`normalizedSubject`](#normalizedsubject-string) 属性获取不带任何前导前缀（如 `RE:` 和 `FW:`）的主题。</span><span class="sxs-lookup"><span data-stu-id="2a739-p131">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="2a739-643">撰写模式</span><span class="sxs-lookup"><span data-stu-id="2a739-643">Compose mode</span></span>

<span data-ttu-id="2a739-644">`subject` 属性返回一个 `Subject` 对象，该对象提供用于获取和设置主题的方法。</span><span class="sxs-lookup"><span data-stu-id="2a739-644">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="2a739-645">类型：</span><span class="sxs-lookup"><span data-stu-id="2a739-645">Type:</span></span>

*   <span data-ttu-id="2a739-646">String | [Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="2a739-646">String | [Subject](/javascript/api/outlook/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="2a739-647">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-647">Requirements</span></span>

|<span data-ttu-id="2a739-648">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-648">Requirement</span></span>|<span data-ttu-id="2a739-649">值</span><span class="sxs-lookup"><span data-stu-id="2a739-649">Value</span></span>|
|---|---|
|[<span data-ttu-id="2a739-650">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2a739-650">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2a739-651">1.0</span><span class="sxs-lookup"><span data-stu-id="2a739-651">1.0</span></span>|
|[<span data-ttu-id="2a739-652">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2a739-652">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2a739-653">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2a739-653">ReadItem</span></span>|
|[<span data-ttu-id="2a739-654">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2a739-654">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2a739-655">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="2a739-655">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="2a739-656">到： 数组。 <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[收件人](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="2a739-656">to :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="2a739-657">提供对邮件的**到**行上的收件人访问。</span><span class="sxs-lookup"><span data-stu-id="2a739-657">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="2a739-658">对象的类型和访问级别取决于当前项的模式。</span><span class="sxs-lookup"><span data-stu-id="2a739-658">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="2a739-659">阅读模式</span><span class="sxs-lookup"><span data-stu-id="2a739-659">Read mode</span></span>

<span data-ttu-id="2a739-p133">`to` 属性返回包含邮件的**收件人**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="2a739-p133">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="2a739-662">撰写模式</span><span class="sxs-lookup"><span data-stu-id="2a739-662">Compose mode</span></span>

<span data-ttu-id="2a739-663">`to`属性返回`Recipients`提供方法用于获取或更新邮件的**到**行上的收件人对象。</span><span class="sxs-lookup"><span data-stu-id="2a739-663">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="2a739-664">类型：</span><span class="sxs-lookup"><span data-stu-id="2a739-664">Type:</span></span>

*   <span data-ttu-id="2a739-665">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="2a739-665">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="2a739-666">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-666">Requirements</span></span>

|<span data-ttu-id="2a739-667">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-667">Requirement</span></span>|<span data-ttu-id="2a739-668">值</span><span class="sxs-lookup"><span data-stu-id="2a739-668">Value</span></span>|
|---|---|
|[<span data-ttu-id="2a739-669">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2a739-669">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2a739-670">1.0</span><span class="sxs-lookup"><span data-stu-id="2a739-670">1.0</span></span>|
|[<span data-ttu-id="2a739-671">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2a739-671">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2a739-672">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2a739-672">ReadItem</span></span>|
|[<span data-ttu-id="2a739-673">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2a739-673">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2a739-674">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="2a739-674">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="2a739-675">示例</span><span class="sxs-lookup"><span data-stu-id="2a739-675">Example</span></span>

```
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="2a739-676">方法</span><span class="sxs-lookup"><span data-stu-id="2a739-676">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="2a739-677">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="2a739-677">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="2a739-678">将文件作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="2a739-678">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="2a739-679">`addFileAttachmentAsync` 方法在指定的 URI 上载文件并将其附加到撰写窗体中的项目。</span><span class="sxs-lookup"><span data-stu-id="2a739-679">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="2a739-680">你随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="2a739-680">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2a739-681">参数：</span><span class="sxs-lookup"><span data-stu-id="2a739-681">Parameters:</span></span>
|<span data-ttu-id="2a739-682">名称</span><span class="sxs-lookup"><span data-stu-id="2a739-682">Name</span></span>|<span data-ttu-id="2a739-683">类型</span><span class="sxs-lookup"><span data-stu-id="2a739-683">Type</span></span>|<span data-ttu-id="2a739-684">属性</span><span class="sxs-lookup"><span data-stu-id="2a739-684">Attributes</span></span>|<span data-ttu-id="2a739-685">说明</span><span class="sxs-lookup"><span data-stu-id="2a739-685">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="2a739-686">字符串</span><span class="sxs-lookup"><span data-stu-id="2a739-686">String</span></span>||<span data-ttu-id="2a739-p134">提供附加到邮件或约会的文件的位置的 URI。最大长度为 2048 个字符。</span><span class="sxs-lookup"><span data-stu-id="2a739-p134">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="2a739-689">字符串</span><span class="sxs-lookup"><span data-stu-id="2a739-689">String</span></span>||<span data-ttu-id="2a739-p135">在附件上载过程中显示的附件名称。最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="2a739-p135">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="2a739-692">对象</span><span class="sxs-lookup"><span data-stu-id="2a739-692">Object</span></span>|<span data-ttu-id="2a739-693">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2a739-693">&lt;optional&gt;</span></span>|<span data-ttu-id="2a739-694">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="2a739-694">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="2a739-695">对象</span><span class="sxs-lookup"><span data-stu-id="2a739-695">Object</span></span>|<span data-ttu-id="2a739-696">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2a739-696">&lt;optional&gt;</span></span>|<span data-ttu-id="2a739-697">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="2a739-697">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="2a739-698">布尔值</span><span class="sxs-lookup"><span data-stu-id="2a739-698">Boolean</span></span>|<span data-ttu-id="2a739-699">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2a739-699">&lt;optional&gt;</span></span>|<span data-ttu-id="2a739-700">如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="2a739-700">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="2a739-701">函数</span><span class="sxs-lookup"><span data-stu-id="2a739-701">function</span></span>|<span data-ttu-id="2a739-702">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2a739-702">&lt;optional&gt;</span></span>|<span data-ttu-id="2a739-703">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="2a739-703">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="2a739-704">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="2a739-704">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="2a739-705">如果上传附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="2a739-705">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="2a739-706">错误</span><span class="sxs-lookup"><span data-stu-id="2a739-706">Errors</span></span>

|<span data-ttu-id="2a739-707">错误代码</span><span class="sxs-lookup"><span data-stu-id="2a739-707">Error code</span></span>|<span data-ttu-id="2a739-708">说明</span><span class="sxs-lookup"><span data-stu-id="2a739-708">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="2a739-709">附件大小超过了允许的大小。</span><span class="sxs-lookup"><span data-stu-id="2a739-709">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="2a739-710">该附件的扩展名不是允许的扩展名。</span><span class="sxs-lookup"><span data-stu-id="2a739-710">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="2a739-711">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="2a739-711">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2a739-712">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-712">Requirements</span></span>

|<span data-ttu-id="2a739-713">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-713">Requirement</span></span>|<span data-ttu-id="2a739-714">值</span><span class="sxs-lookup"><span data-stu-id="2a739-714">Value</span></span>|
|---|---|
|[<span data-ttu-id="2a739-715">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2a739-715">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2a739-716">1.1</span><span class="sxs-lookup"><span data-stu-id="2a739-716">1.1</span></span>|
|[<span data-ttu-id="2a739-717">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2a739-717">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2a739-718">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="2a739-718">ReadWriteItem</span></span>|
|[<span data-ttu-id="2a739-719">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2a739-719">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2a739-720">撰写</span><span class="sxs-lookup"><span data-stu-id="2a739-720">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="2a739-721">示例</span><span class="sxs-lookup"><span data-stu-id="2a739-721">Examples</span></span>

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

<span data-ttu-id="2a739-722">以下示例将图像文件添加为内联附件，并在邮件正文中引用该附件。</span><span class="sxs-lookup"><span data-stu-id="2a739-722">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a><span data-ttu-id="2a739-723">addFileAttachmentFromBase64Async (base64File，attachmentName，[选项] [回调])</span><span class="sxs-lookup"><span data-stu-id="2a739-723">addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="2a739-724">添加从 base64 编码到邮件或约会作为附件的文件。</span><span class="sxs-lookup"><span data-stu-id="2a739-724">Adds a file from the base64 encoding to a message or appointment as an attachment.</span></span>

<span data-ttu-id="2a739-725">`addFileAttachmentFromBase64Async`方法从 base64 编码将文件上载，并将其附加到撰写窗体中的项。</span><span class="sxs-lookup"><span data-stu-id="2a739-725">The `addFileAttachmentFromBase64Async` method uploads the file from the base64 encoding and attaches it to the item in the compose form.</span></span> <span data-ttu-id="2a739-726">此方法返回 AsyncResult.value 对象中的附件标识符。</span><span class="sxs-lookup"><span data-stu-id="2a739-726">This method returns the attachment identifier in the AsyncResult.value object.</span></span>

<span data-ttu-id="2a739-727">你随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="2a739-727">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2a739-728">参数：</span><span class="sxs-lookup"><span data-stu-id="2a739-728">Parameters:</span></span>
|<span data-ttu-id="2a739-729">名称</span><span class="sxs-lookup"><span data-stu-id="2a739-729">Name</span></span>|<span data-ttu-id="2a739-730">类型</span><span class="sxs-lookup"><span data-stu-id="2a739-730">Type</span></span>|<span data-ttu-id="2a739-731">属性</span><span class="sxs-lookup"><span data-stu-id="2a739-731">Attributes</span></span>|<span data-ttu-id="2a739-732">说明</span><span class="sxs-lookup"><span data-stu-id="2a739-732">Description</span></span>|
|---|---|---|---|
|`base64File`|<span data-ttu-id="2a739-733">String</span><span class="sxs-lookup"><span data-stu-id="2a739-733">String</span></span>||<span data-ttu-id="2a739-734">以 base64 编码的图像或文件添加到电子邮件或事件的内容。</span><span class="sxs-lookup"><span data-stu-id="2a739-734">The base64 encoded content of an image or file to be added to an email or event.</span></span>|
|`attachmentName`|<span data-ttu-id="2a739-735">字符串</span><span class="sxs-lookup"><span data-stu-id="2a739-735">String</span></span>||<span data-ttu-id="2a739-p137">在附件上载过程中显示的附件名称。最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="2a739-p137">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="2a739-738">对象</span><span class="sxs-lookup"><span data-stu-id="2a739-738">Object</span></span>|<span data-ttu-id="2a739-739">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2a739-739">&lt;optional&gt;</span></span>|<span data-ttu-id="2a739-740">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="2a739-740">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="2a739-741">对象</span><span class="sxs-lookup"><span data-stu-id="2a739-741">Object</span></span>|<span data-ttu-id="2a739-742">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2a739-742">&lt;optional&gt;</span></span>|<span data-ttu-id="2a739-743">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="2a739-743">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="2a739-744">布尔值</span><span class="sxs-lookup"><span data-stu-id="2a739-744">Boolean</span></span>|<span data-ttu-id="2a739-745">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2a739-745">&lt;optional&gt;</span></span>|<span data-ttu-id="2a739-746">如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="2a739-746">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="2a739-747">函数</span><span class="sxs-lookup"><span data-stu-id="2a739-747">function</span></span>|<span data-ttu-id="2a739-748">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2a739-748">&lt;optional&gt;</span></span>|<span data-ttu-id="2a739-749">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="2a739-749">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="2a739-750">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="2a739-750">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="2a739-751">如果上传附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="2a739-751">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="2a739-752">错误</span><span class="sxs-lookup"><span data-stu-id="2a739-752">Errors</span></span>

|<span data-ttu-id="2a739-753">错误代码</span><span class="sxs-lookup"><span data-stu-id="2a739-753">Error code</span></span>|<span data-ttu-id="2a739-754">说明</span><span class="sxs-lookup"><span data-stu-id="2a739-754">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="2a739-755">附件大小超过了允许的大小。</span><span class="sxs-lookup"><span data-stu-id="2a739-755">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="2a739-756">该附件的扩展名不是允许的扩展名。</span><span class="sxs-lookup"><span data-stu-id="2a739-756">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="2a739-757">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="2a739-757">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2a739-758">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-758">Requirements</span></span>

|<span data-ttu-id="2a739-759">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-759">Requirement</span></span>|<span data-ttu-id="2a739-760">值</span><span class="sxs-lookup"><span data-stu-id="2a739-760">Value</span></span>|
|---|---|
|[<span data-ttu-id="2a739-761">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2a739-761">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2a739-762">预览</span><span class="sxs-lookup"><span data-stu-id="2a739-762">Preview</span></span>|
|[<span data-ttu-id="2a739-763">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2a739-763">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2a739-764">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="2a739-764">ReadWriteItem</span></span>|
|[<span data-ttu-id="2a739-765">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2a739-765">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2a739-766">撰写</span><span class="sxs-lookup"><span data-stu-id="2a739-766">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="2a739-767">示例</span><span class="sxs-lookup"><span data-stu-id="2a739-767">Examples</span></span>

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

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="2a739-768">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="2a739-768">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="2a739-769">添加支持事件的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="2a739-769">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="2a739-770">当前受支持的事件类型的`Office.EventType.AppointmentTimeChanged`， `Office.EventType.RecipientsChanged`，和`Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="2a739-770">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="2a739-771">参数：</span><span class="sxs-lookup"><span data-stu-id="2a739-771">Parameters:</span></span>

| <span data-ttu-id="2a739-772">名称</span><span class="sxs-lookup"><span data-stu-id="2a739-772">Name</span></span> | <span data-ttu-id="2a739-773">类型</span><span class="sxs-lookup"><span data-stu-id="2a739-773">Type</span></span> | <span data-ttu-id="2a739-774">属性</span><span class="sxs-lookup"><span data-stu-id="2a739-774">Attributes</span></span> | <span data-ttu-id="2a739-775">说明</span><span class="sxs-lookup"><span data-stu-id="2a739-775">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="2a739-776">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="2a739-776">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="2a739-777">应调用处理程序的事件。</span><span class="sxs-lookup"><span data-stu-id="2a739-777">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="2a739-778">函数</span><span class="sxs-lookup"><span data-stu-id="2a739-778">Function</span></span> || <span data-ttu-id="2a739-p138">用于处理事件的函数。此函数必须接受一个参数，即对象文本。参数上的 `type` 属性将匹配传递给 `addHandlerAsync` 的 `eventType` 参数。</span><span class="sxs-lookup"><span data-stu-id="2a739-p138">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="2a739-782">Object</span><span class="sxs-lookup"><span data-stu-id="2a739-782">Object</span></span> | <span data-ttu-id="2a739-783">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2a739-783">&lt;optional&gt;</span></span> | <span data-ttu-id="2a739-784">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="2a739-784">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="2a739-785">对象</span><span class="sxs-lookup"><span data-stu-id="2a739-785">Object</span></span> | <span data-ttu-id="2a739-786">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2a739-786">&lt;optional&gt;</span></span> | <span data-ttu-id="2a739-787">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="2a739-787">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="2a739-788">函数</span><span class="sxs-lookup"><span data-stu-id="2a739-788">function</span></span>| <span data-ttu-id="2a739-789">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2a739-789">&lt;optional&gt;</span></span>|<span data-ttu-id="2a739-790">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="2a739-790">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2a739-791">Requirements</span><span class="sxs-lookup"><span data-stu-id="2a739-791">Requirements</span></span>

|<span data-ttu-id="2a739-792">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-792">Requirement</span></span>| <span data-ttu-id="2a739-793">值</span><span class="sxs-lookup"><span data-stu-id="2a739-793">Value</span></span>|
|---|---|
|[<span data-ttu-id="2a739-794">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2a739-794">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2a739-795">1.7</span><span class="sxs-lookup"><span data-stu-id="2a739-795">1.7</span></span> |
|[<span data-ttu-id="2a739-796">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2a739-796">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2a739-797">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2a739-797">ReadItem</span></span> |
|[<span data-ttu-id="2a739-798">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2a739-798">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2a739-799">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="2a739-799">Compose or read</span></span> |

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="2a739-800">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="2a739-800">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="2a739-801">将 Exchange 项目（如邮件）作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="2a739-801">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="2a739-p139">`addItemAttachmentAsync` 方法将包含指定 Exchange 标识符的项目附加到撰写窗体中的项目。如果指定一个回调方法，此方法使用 `asyncResult` 参数调用，该参数包含一个附件标识符或代码，指示附加项目过程中出现的任何错误。可以使用 `options` 参数将状态信息传递给回调方法（如果需要）。</span><span class="sxs-lookup"><span data-stu-id="2a739-p139">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="2a739-805">你随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="2a739-805">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="2a739-806">如果您 Office 加载项运行的 Outlook Web App 中`addItemAttachmentAsync`方法可以附加到正在编辑; 的项目之外的项目的项目但是，这不受支持，并且不建议使用。</span><span class="sxs-lookup"><span data-stu-id="2a739-806">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2a739-807">参数：</span><span class="sxs-lookup"><span data-stu-id="2a739-807">Parameters:</span></span>

|<span data-ttu-id="2a739-808">名称</span><span class="sxs-lookup"><span data-stu-id="2a739-808">Name</span></span>|<span data-ttu-id="2a739-809">类型</span><span class="sxs-lookup"><span data-stu-id="2a739-809">Type</span></span>|<span data-ttu-id="2a739-810">属性</span><span class="sxs-lookup"><span data-stu-id="2a739-810">Attributes</span></span>|<span data-ttu-id="2a739-811">说明</span><span class="sxs-lookup"><span data-stu-id="2a739-811">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="2a739-812">字符串</span><span class="sxs-lookup"><span data-stu-id="2a739-812">String</span></span>||<span data-ttu-id="2a739-p140">要附加的项目的 Exchange 标识符。最大长度为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="2a739-p140">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="2a739-815">字符串</span><span class="sxs-lookup"><span data-stu-id="2a739-815">String</span></span>||<span data-ttu-id="2a739-p141">要附加的项目的主题。最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="2a739-p141">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="2a739-818">对象</span><span class="sxs-lookup"><span data-stu-id="2a739-818">Object</span></span>|<span data-ttu-id="2a739-819">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2a739-819">&lt;optional&gt;</span></span>|<span data-ttu-id="2a739-820">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="2a739-820">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="2a739-821">对象</span><span class="sxs-lookup"><span data-stu-id="2a739-821">Object</span></span>|<span data-ttu-id="2a739-822">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2a739-822">&lt;optional&gt;</span></span>|<span data-ttu-id="2a739-823">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="2a739-823">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="2a739-824">函数</span><span class="sxs-lookup"><span data-stu-id="2a739-824">function</span></span>|<span data-ttu-id="2a739-825">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2a739-825">&lt;optional&gt;</span></span>|<span data-ttu-id="2a739-826">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="2a739-826">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="2a739-827">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="2a739-827">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="2a739-828">如果添加附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="2a739-828">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="2a739-829">错误</span><span class="sxs-lookup"><span data-stu-id="2a739-829">Errors</span></span>

|<span data-ttu-id="2a739-830">错误代码</span><span class="sxs-lookup"><span data-stu-id="2a739-830">Error code</span></span>|<span data-ttu-id="2a739-831">说明</span><span class="sxs-lookup"><span data-stu-id="2a739-831">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="2a739-832">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="2a739-832">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2a739-833">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-833">Requirements</span></span>

|<span data-ttu-id="2a739-834">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-834">Requirement</span></span>|<span data-ttu-id="2a739-835">值</span><span class="sxs-lookup"><span data-stu-id="2a739-835">Value</span></span>|
|---|---|
|[<span data-ttu-id="2a739-836">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2a739-836">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2a739-837">1.1</span><span class="sxs-lookup"><span data-stu-id="2a739-837">1.1</span></span>|
|[<span data-ttu-id="2a739-838">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2a739-838">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2a739-839">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="2a739-839">ReadWriteItem</span></span>|
|[<span data-ttu-id="2a739-840">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2a739-840">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2a739-841">撰写</span><span class="sxs-lookup"><span data-stu-id="2a739-841">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="2a739-842">示例</span><span class="sxs-lookup"><span data-stu-id="2a739-842">Example</span></span>

<span data-ttu-id="2a739-843">以下示例将现有的 Outlook 项目添加为名为 `My Attachment` 的附件。</span><span class="sxs-lookup"><span data-stu-id="2a739-843">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

####  <a name="close"></a><span data-ttu-id="2a739-844">close()</span><span class="sxs-lookup"><span data-stu-id="2a739-844">close()</span></span>

<span data-ttu-id="2a739-845">关闭当前正在撰写的项目。</span><span class="sxs-lookup"><span data-stu-id="2a739-845">Closes the current item that is being composed.</span></span>

<span data-ttu-id="2a739-p142">`close` 方法的行为取决于要撰写的项目的当前状态。如果项目具有未保存的更改，客户端将提示用户保存、放弃或取消关闭操作。</span><span class="sxs-lookup"><span data-stu-id="2a739-p142">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="2a739-848">在 Outlook 中的 web，如果该项目是约会和它之前已保存使用在`saveAsync`、 提示用户保存、 放弃，或取消，即使未发生更改由于项目上次保存。</span><span class="sxs-lookup"><span data-stu-id="2a739-848">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="2a739-849">在 Outlook 桌面客户端中，如果邮件是内联答复，`close` 方法不起作用。</span><span class="sxs-lookup"><span data-stu-id="2a739-849">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="2a739-850">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-850">Requirements</span></span>

|<span data-ttu-id="2a739-851">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-851">Requirement</span></span>|<span data-ttu-id="2a739-852">值</span><span class="sxs-lookup"><span data-stu-id="2a739-852">Value</span></span>|
|---|---|
|[<span data-ttu-id="2a739-853">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2a739-853">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2a739-854">1.3</span><span class="sxs-lookup"><span data-stu-id="2a739-854">1.3</span></span>|
|[<span data-ttu-id="2a739-855">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2a739-855">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2a739-856">受限</span><span class="sxs-lookup"><span data-stu-id="2a739-856">Restricted</span></span>|
|[<span data-ttu-id="2a739-857">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2a739-857">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2a739-858">撰写</span><span class="sxs-lookup"><span data-stu-id="2a739-858">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="2a739-859">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="2a739-859">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="2a739-860">显示答复窗体，其中包括所选邮件的发件人和所有收件人或所选约会的组织者和所有与会者。</span><span class="sxs-lookup"><span data-stu-id="2a739-860">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="2a739-861">在适用于 iOS 的 Outlook 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="2a739-861">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="2a739-862">在 Outlook Web App 中，答复窗体显示为包含 3 列视图的弹出式窗体以及包含 2 列或 1 列视图的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="2a739-862">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="2a739-863">如果任意字符串参数超出其限制，`displayReplyAllForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="2a739-863">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="2a739-p143">当在 `formData.attachments` 参数中指定附件时，Outlook 和 Outlook Web App 尝试下载所有附件并将其附加到答复窗体。如果无法添加任何附件，则在窗体 UI 中显示错误。如果这不可能，则不引发错误消息。</span><span class="sxs-lookup"><span data-stu-id="2a739-p143">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2a739-867">参数：</span><span class="sxs-lookup"><span data-stu-id="2a739-867">Parameters:</span></span>

|<span data-ttu-id="2a739-868">名称</span><span class="sxs-lookup"><span data-stu-id="2a739-868">Name</span></span>|<span data-ttu-id="2a739-869">类型</span><span class="sxs-lookup"><span data-stu-id="2a739-869">Type</span></span>|<span data-ttu-id="2a739-870">属性</span><span class="sxs-lookup"><span data-stu-id="2a739-870">Attributes</span></span>|<span data-ttu-id="2a739-871">说明</span><span class="sxs-lookup"><span data-stu-id="2a739-871">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="2a739-872">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="2a739-872">String &#124; Object</span></span>||<span data-ttu-id="2a739-p144">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="2a739-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="2a739-875">**或**</span><span class="sxs-lookup"><span data-stu-id="2a739-875">**OR**</span></span><br/><span data-ttu-id="2a739-p145">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="2a739-p145">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="2a739-878">字符串</span><span class="sxs-lookup"><span data-stu-id="2a739-878">String</span></span>|<span data-ttu-id="2a739-879">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2a739-879">&lt;optional&gt;</span></span>|<span data-ttu-id="2a739-p146">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="2a739-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="2a739-882">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="2a739-882">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="2a739-883">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2a739-883">&lt;optional&gt;</span></span>|<span data-ttu-id="2a739-884">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="2a739-884">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="2a739-885">字符串</span><span class="sxs-lookup"><span data-stu-id="2a739-885">String</span></span>||<span data-ttu-id="2a739-p147">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="2a739-p147">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="2a739-888">字符串</span><span class="sxs-lookup"><span data-stu-id="2a739-888">String</span></span>||<span data-ttu-id="2a739-889">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="2a739-889">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="2a739-890">String</span><span class="sxs-lookup"><span data-stu-id="2a739-890">String</span></span>||<span data-ttu-id="2a739-p148">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="2a739-p148">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="2a739-893">Boolean</span><span class="sxs-lookup"><span data-stu-id="2a739-893">Boolean</span></span>||<span data-ttu-id="2a739-p149">仅在将 `type` 设置为 `file` 时使用。如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="2a739-p149">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="2a739-896">String</span><span class="sxs-lookup"><span data-stu-id="2a739-896">String</span></span>||<span data-ttu-id="2a739-p150">仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="2a739-p150">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="2a739-900">函数</span><span class="sxs-lookup"><span data-stu-id="2a739-900">function</span></span>|<span data-ttu-id="2a739-901">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2a739-901">&lt;optional&gt;</span></span>|<span data-ttu-id="2a739-902">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="2a739-902">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2a739-903">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-903">Requirements</span></span>

|<span data-ttu-id="2a739-904">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-904">Requirement</span></span>|<span data-ttu-id="2a739-905">值</span><span class="sxs-lookup"><span data-stu-id="2a739-905">Value</span></span>|
|---|---|
|[<span data-ttu-id="2a739-906">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2a739-906">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2a739-907">1.0</span><span class="sxs-lookup"><span data-stu-id="2a739-907">1.0</span></span>|
|[<span data-ttu-id="2a739-908">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2a739-908">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2a739-909">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2a739-909">ReadItem</span></span>|
|[<span data-ttu-id="2a739-910">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2a739-910">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2a739-911">阅读</span><span class="sxs-lookup"><span data-stu-id="2a739-911">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="2a739-912">示例</span><span class="sxs-lookup"><span data-stu-id="2a739-912">Examples</span></span>

<span data-ttu-id="2a739-913">以下代码将一个字符串传递到 `displayReplyAllForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="2a739-913">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="2a739-914">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="2a739-914">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="2a739-915">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="2a739-915">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="2a739-916">使用正文和文件附件答复。</span><span class="sxs-lookup"><span data-stu-id="2a739-916">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="2a739-917">使用正文和项目附件答复。</span><span class="sxs-lookup"><span data-stu-id="2a739-917">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="2a739-918">使用正文、文件附件、项目附件和回调答复。</span><span class="sxs-lookup"><span data-stu-id="2a739-918">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="2a739-919">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="2a739-919">displayReplyForm(formData)</span></span>

<span data-ttu-id="2a739-920">显示答复窗体，其中仅包括所选邮件的发件人或所选约会的组织者。</span><span class="sxs-lookup"><span data-stu-id="2a739-920">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="2a739-921">在适用于 iOS 的 Outlook 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="2a739-921">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="2a739-922">在 Outlook Web App 中，答复窗体显示为包含 3 列视图的弹出式窗体以及包含 2 列或 1 列视图的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="2a739-922">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="2a739-923">如果任意字符串参数超出其限制，`displayReplyForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="2a739-923">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="2a739-p151">当在 `formData.attachments` 参数中指定附件时，Outlook 和 Outlook Web App 尝试下载所有附件并将其附加到答复窗体。如果无法添加任何附件，则在窗体 UI 中显示错误。如果这不可能，则不引发错误消息。</span><span class="sxs-lookup"><span data-stu-id="2a739-p151">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2a739-927">参数：</span><span class="sxs-lookup"><span data-stu-id="2a739-927">Parameters:</span></span>

|<span data-ttu-id="2a739-928">名称</span><span class="sxs-lookup"><span data-stu-id="2a739-928">Name</span></span>|<span data-ttu-id="2a739-929">类型</span><span class="sxs-lookup"><span data-stu-id="2a739-929">Type</span></span>|<span data-ttu-id="2a739-930">属性</span><span class="sxs-lookup"><span data-stu-id="2a739-930">Attributes</span></span>|<span data-ttu-id="2a739-931">说明</span><span class="sxs-lookup"><span data-stu-id="2a739-931">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="2a739-932">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="2a739-932">String &#124; Object</span></span>||<span data-ttu-id="2a739-p152">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="2a739-p152">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="2a739-935">**或**</span><span class="sxs-lookup"><span data-stu-id="2a739-935">**OR**</span></span><br/><span data-ttu-id="2a739-p153">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="2a739-p153">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="2a739-938">字符串</span><span class="sxs-lookup"><span data-stu-id="2a739-938">String</span></span>|<span data-ttu-id="2a739-939">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2a739-939">&lt;optional&gt;</span></span>|<span data-ttu-id="2a739-p154">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="2a739-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="2a739-942">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="2a739-942">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="2a739-943">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2a739-943">&lt;optional&gt;</span></span>|<span data-ttu-id="2a739-944">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="2a739-944">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="2a739-945">字符串</span><span class="sxs-lookup"><span data-stu-id="2a739-945">String</span></span>||<span data-ttu-id="2a739-p155">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="2a739-p155">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="2a739-948">字符串</span><span class="sxs-lookup"><span data-stu-id="2a739-948">String</span></span>||<span data-ttu-id="2a739-949">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="2a739-949">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="2a739-950">String</span><span class="sxs-lookup"><span data-stu-id="2a739-950">String</span></span>||<span data-ttu-id="2a739-p156">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="2a739-p156">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="2a739-953">Boolean</span><span class="sxs-lookup"><span data-stu-id="2a739-953">Boolean</span></span>||<span data-ttu-id="2a739-p157">仅在将 `type` 设置为 `file` 时使用。如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="2a739-p157">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="2a739-956">String</span><span class="sxs-lookup"><span data-stu-id="2a739-956">String</span></span>||<span data-ttu-id="2a739-p158">仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="2a739-p158">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="2a739-960">函数</span><span class="sxs-lookup"><span data-stu-id="2a739-960">function</span></span>|<span data-ttu-id="2a739-961">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2a739-961">&lt;optional&gt;</span></span>|<span data-ttu-id="2a739-962">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="2a739-962">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2a739-963">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-963">Requirements</span></span>

|<span data-ttu-id="2a739-964">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-964">Requirement</span></span>|<span data-ttu-id="2a739-965">值</span><span class="sxs-lookup"><span data-stu-id="2a739-965">Value</span></span>|
|---|---|
|[<span data-ttu-id="2a739-966">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2a739-966">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2a739-967">1.0</span><span class="sxs-lookup"><span data-stu-id="2a739-967">1.0</span></span>|
|[<span data-ttu-id="2a739-968">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2a739-968">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2a739-969">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2a739-969">ReadItem</span></span>|
|[<span data-ttu-id="2a739-970">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2a739-970">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2a739-971">阅读</span><span class="sxs-lookup"><span data-stu-id="2a739-971">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="2a739-972">示例</span><span class="sxs-lookup"><span data-stu-id="2a739-972">Examples</span></span>

<span data-ttu-id="2a739-973">以下代码将一个字符串传递到 `displayReplyForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="2a739-973">The following code passes a string to the `displayReplyForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="2a739-974">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="2a739-974">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="2a739-975">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="2a739-975">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="2a739-976">使用正文和文件附件答复。</span><span class="sxs-lookup"><span data-stu-id="2a739-976">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="2a739-977">使用正文和项目附件答复。</span><span class="sxs-lookup"><span data-stu-id="2a739-977">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="2a739-978">使用正文、文件附件、项目附件和回调答复。</span><span class="sxs-lookup"><span data-stu-id="2a739-978">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="2a739-979">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="2a739-979">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="2a739-980">获取在选定的项的正文中找到的实体。</span><span class="sxs-lookup"><span data-stu-id="2a739-980">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="2a739-981">在适用于 iOS 的 Outlook 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="2a739-981">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="2a739-982">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-982">Requirements</span></span>

|<span data-ttu-id="2a739-983">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-983">Requirement</span></span>|<span data-ttu-id="2a739-984">值</span><span class="sxs-lookup"><span data-stu-id="2a739-984">Value</span></span>|
|---|---|
|[<span data-ttu-id="2a739-985">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2a739-985">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2a739-986">1.0</span><span class="sxs-lookup"><span data-stu-id="2a739-986">1.0</span></span>|
|[<span data-ttu-id="2a739-987">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2a739-987">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2a739-988">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2a739-988">ReadItem</span></span>|
|[<span data-ttu-id="2a739-989">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2a739-989">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2a739-990">阅读</span><span class="sxs-lookup"><span data-stu-id="2a739-990">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="2a739-991">返回：</span><span class="sxs-lookup"><span data-stu-id="2a739-991">Returns:</span></span>

<span data-ttu-id="2a739-992">类型：[Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="2a739-992">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="2a739-993">示例</span><span class="sxs-lookup"><span data-stu-id="2a739-993">Example</span></span>

<span data-ttu-id="2a739-994">以下示例访问当前项的主体中的联系人实体。</span><span class="sxs-lookup"><span data-stu-id="2a739-994">The following example accesses the contacts entities in the current item's body.</span></span>

```
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="2a739-995">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="2a739-995">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="2a739-996">获取选定的项的正文中找到的指定的实体类型的所有实体的数组。</span><span class="sxs-lookup"><span data-stu-id="2a739-996">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="2a739-997">在适用于 iOS 的 Outlook 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="2a739-997">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2a739-998">参数：</span><span class="sxs-lookup"><span data-stu-id="2a739-998">Parameters:</span></span>

|<span data-ttu-id="2a739-999">名称</span><span class="sxs-lookup"><span data-stu-id="2a739-999">Name</span></span>|<span data-ttu-id="2a739-1000">类型</span><span class="sxs-lookup"><span data-stu-id="2a739-1000">Type</span></span>|<span data-ttu-id="2a739-1001">说明</span><span class="sxs-lookup"><span data-stu-id="2a739-1001">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="2a739-1002">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="2a739-1002">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype)|<span data-ttu-id="2a739-1003">EntityType 枚举值之一。</span><span class="sxs-lookup"><span data-stu-id="2a739-1003">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2a739-1004">Requirements</span><span class="sxs-lookup"><span data-stu-id="2a739-1004">Requirements</span></span>

|<span data-ttu-id="2a739-1005">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-1005">Requirement</span></span>|<span data-ttu-id="2a739-1006">值</span><span class="sxs-lookup"><span data-stu-id="2a739-1006">Value</span></span>|
|---|---|
|[<span data-ttu-id="2a739-1007">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2a739-1007">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2a739-1008">1.0</span><span class="sxs-lookup"><span data-stu-id="2a739-1008">1.0</span></span>|
|[<span data-ttu-id="2a739-1009">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2a739-1009">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2a739-1010">受限</span><span class="sxs-lookup"><span data-stu-id="2a739-1010">Restricted</span></span>|
|[<span data-ttu-id="2a739-1011">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2a739-1011">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2a739-1012">阅读</span><span class="sxs-lookup"><span data-stu-id="2a739-1012">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="2a739-1013">返回：</span><span class="sxs-lookup"><span data-stu-id="2a739-1013">Returns:</span></span>

<span data-ttu-id="2a739-1014">如果在 `entityType` 中传递的值不是 `EntityType` 枚举的有效成员，该方法返回 null。</span><span class="sxs-lookup"><span data-stu-id="2a739-1014">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="2a739-1015">如果没有指定类型的实体的存在于项目的正文中，该方法将返回空数组。</span><span class="sxs-lookup"><span data-stu-id="2a739-1015">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="2a739-1016">否则，返回的数组中对象的类型取决于 `entityType` 参数中请求实体的类型。</span><span class="sxs-lookup"><span data-stu-id="2a739-1016">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="2a739-1017">当使用此方法的最低权限级别**受限**时，某些实体类型需要 **ReadItem** 才能进行访问，如下表中所指定。</span><span class="sxs-lookup"><span data-stu-id="2a739-1017">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="2a739-1018">`entityType` 的值</span><span class="sxs-lookup"><span data-stu-id="2a739-1018">Value of `entityType`</span></span>|<span data-ttu-id="2a739-1019">返回的数组中对象的类型</span><span class="sxs-lookup"><span data-stu-id="2a739-1019">Type of objects in returned array</span></span>|<span data-ttu-id="2a739-1020">所需权限级别</span><span class="sxs-lookup"><span data-stu-id="2a739-1020">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="2a739-1021">String</span><span class="sxs-lookup"><span data-stu-id="2a739-1021">String</span></span>|<span data-ttu-id="2a739-1022">**受限**</span><span class="sxs-lookup"><span data-stu-id="2a739-1022">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="2a739-1023">Contact</span><span class="sxs-lookup"><span data-stu-id="2a739-1023">Contact</span></span>|<span data-ttu-id="2a739-1024">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="2a739-1024">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="2a739-1025">String</span><span class="sxs-lookup"><span data-stu-id="2a739-1025">String</span></span>|<span data-ttu-id="2a739-1026">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="2a739-1026">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="2a739-1027">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="2a739-1027">MeetingSuggestion</span></span>|<span data-ttu-id="2a739-1028">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="2a739-1028">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="2a739-1029">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="2a739-1029">PhoneNumber</span></span>|<span data-ttu-id="2a739-1030">**受限**</span><span class="sxs-lookup"><span data-stu-id="2a739-1030">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="2a739-1031">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="2a739-1031">TaskSuggestion</span></span>|<span data-ttu-id="2a739-1032">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="2a739-1032">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="2a739-1033">String</span><span class="sxs-lookup"><span data-stu-id="2a739-1033">String</span></span>|<span data-ttu-id="2a739-1034">**受限**</span><span class="sxs-lookup"><span data-stu-id="2a739-1034">**Restricted**</span></span>|

<span data-ttu-id="2a739-1035">类型：Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="2a739-1035">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="2a739-1036">示例</span><span class="sxs-lookup"><span data-stu-id="2a739-1036">Example</span></span>

<span data-ttu-id="2a739-1037">下面的示例演示如何访问表示当前项目的正文中的邮政地址的字符串数组。</span><span class="sxs-lookup"><span data-stu-id="2a739-1037">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="2a739-1038">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="2a739-1038">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="2a739-1039">返回传递清单 XML 文件中定义的命名筛选器的所选项目中的已知实体。</span><span class="sxs-lookup"><span data-stu-id="2a739-1039">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="2a739-1040">在适用于 iOS 的 Outlook 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="2a739-1040">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="2a739-1041">`getFilteredEntitiesByName` 方法返回匹配在具有指定 `FilterName` 元素值的清单 XML 文件中的 [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) 规则元素中定义的正则表达式的实体。</span><span class="sxs-lookup"><span data-stu-id="2a739-1041">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2a739-1042">参数：</span><span class="sxs-lookup"><span data-stu-id="2a739-1042">Parameters:</span></span>

|<span data-ttu-id="2a739-1043">名称</span><span class="sxs-lookup"><span data-stu-id="2a739-1043">Name</span></span>|<span data-ttu-id="2a739-1044">类型</span><span class="sxs-lookup"><span data-stu-id="2a739-1044">Type</span></span>|<span data-ttu-id="2a739-1045">说明</span><span class="sxs-lookup"><span data-stu-id="2a739-1045">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="2a739-1046">字符串</span><span class="sxs-lookup"><span data-stu-id="2a739-1046">String</span></span>|<span data-ttu-id="2a739-1047">定义筛选器匹配的 `ItemHasKnownEntity` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="2a739-1047">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2a739-1048">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-1048">Requirements</span></span>

|<span data-ttu-id="2a739-1049">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-1049">Requirement</span></span>|<span data-ttu-id="2a739-1050">值</span><span class="sxs-lookup"><span data-stu-id="2a739-1050">Value</span></span>|
|---|---|
|[<span data-ttu-id="2a739-1051">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2a739-1051">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2a739-1052">1.0</span><span class="sxs-lookup"><span data-stu-id="2a739-1052">1.0</span></span>|
|[<span data-ttu-id="2a739-1053">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2a739-1053">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2a739-1054">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2a739-1054">ReadItem</span></span>|
|[<span data-ttu-id="2a739-1055">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2a739-1055">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2a739-1056">阅读</span><span class="sxs-lookup"><span data-stu-id="2a739-1056">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="2a739-1057">返回：</span><span class="sxs-lookup"><span data-stu-id="2a739-1057">Returns:</span></span>

<span data-ttu-id="2a739-p160">如果具有匹配 `name` 参数的 `FilterName` 元素值的清单中没有任何 `ItemHasKnownEntity` 元素，则该方法返回 `null`。如果 `name` 参数匹配清单中的 `ItemHasKnownEntity` 元素，但在匹配的当前项目中没有实体，则该方法返回一个空数组。</span><span class="sxs-lookup"><span data-stu-id="2a739-p160">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="2a739-1060">类型：Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="2a739-1060">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

#### <a name="getinitializationcontextasyncoptions-callback"></a><span data-ttu-id="2a739-1061">getInitializationContextAsync([options], [callback])</span><span class="sxs-lookup"><span data-stu-id="2a739-1061">getInitializationContextAsync([options], [callback])</span></span>

<span data-ttu-id="2a739-1062">当[可操作邮件激活](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message)加载项时，获取传递的初始化数据。</span><span class="sxs-lookup"><span data-stu-id="2a739-1062">Gets initialization data passed when the add-in is [activated by an actionable message](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

> [!NOTE]
> <span data-ttu-id="2a739-1063">仅支持使用此方法通过 Outlook 2016 Windows （单击即点即用版本大于 16.0.8413.1000） 和 Outlook web 上的 Office 365。</span><span class="sxs-lookup"><span data-stu-id="2a739-1063">This method is only supported by Outlook 2016 for Windows (Click-to-Run versions greater than 16.0.8413.1000) and Outlook on the web for Office 365.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2a739-1064">参数：</span><span class="sxs-lookup"><span data-stu-id="2a739-1064">Parameters:</span></span>
|<span data-ttu-id="2a739-1065">名称</span><span class="sxs-lookup"><span data-stu-id="2a739-1065">Name</span></span>|<span data-ttu-id="2a739-1066">类型</span><span class="sxs-lookup"><span data-stu-id="2a739-1066">Type</span></span>|<span data-ttu-id="2a739-1067">属性</span><span class="sxs-lookup"><span data-stu-id="2a739-1067">Attributes</span></span>|<span data-ttu-id="2a739-1068">说明</span><span class="sxs-lookup"><span data-stu-id="2a739-1068">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="2a739-1069">对象</span><span class="sxs-lookup"><span data-stu-id="2a739-1069">Object</span></span>|<span data-ttu-id="2a739-1070">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2a739-1070">&lt;optional&gt;</span></span>|<span data-ttu-id="2a739-1071">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="2a739-1071">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="2a739-1072">对象</span><span class="sxs-lookup"><span data-stu-id="2a739-1072">Object</span></span>|<span data-ttu-id="2a739-1073">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2a739-1073">&lt;optional&gt;</span></span>|<span data-ttu-id="2a739-1074">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="2a739-1074">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="2a739-1075">函数</span><span class="sxs-lookup"><span data-stu-id="2a739-1075">function</span></span>|<span data-ttu-id="2a739-1076">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2a739-1076">&lt;optional&gt;</span></span>|<span data-ttu-id="2a739-1077">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="2a739-1077">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="2a739-1078">成功，在初始化数据提供在`asyncResult.value`属性的字符串形式。</span><span class="sxs-lookup"><span data-stu-id="2a739-1078">On success, the initialization data is provided in the `asyncResult.value` property as a string.</span></span><br/><span data-ttu-id="2a739-1079">如果没有初始化上下文，`asyncResult` 对象包含 `Error` 对象，并将它的 `code` 和 `name` 属性分别设置为 `9020` 和 `GenericResponseError`。</span><span class="sxs-lookup"><span data-stu-id="2a739-1079">If there is no initialization context, the `asyncResult` object will contain an `Error` object with its `code` property set to `9020` and its `name` property set to `GenericResponseError`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2a739-1080">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-1080">Requirements</span></span>

|<span data-ttu-id="2a739-1081">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-1081">Requirement</span></span>|<span data-ttu-id="2a739-1082">值</span><span class="sxs-lookup"><span data-stu-id="2a739-1082">Value</span></span>|
|---|---|
|[<span data-ttu-id="2a739-1083">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2a739-1083">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2a739-1084">预览</span><span class="sxs-lookup"><span data-stu-id="2a739-1084">Preview</span></span>|
|[<span data-ttu-id="2a739-1085">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2a739-1085">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2a739-1086">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2a739-1086">ReadItem</span></span>|
|[<span data-ttu-id="2a739-1087">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2a739-1087">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2a739-1088">阅读</span><span class="sxs-lookup"><span data-stu-id="2a739-1088">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2a739-1089">示例</span><span class="sxs-lookup"><span data-stu-id="2a739-1089">Example</span></span>

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

#### <a name="getregexmatches--object"></a><span data-ttu-id="2a739-1090">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="2a739-1090">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="2a739-1091">返回所选项目中匹配在清单 XML 文件中定义的正则表达式的字符串值。</span><span class="sxs-lookup"><span data-stu-id="2a739-1091">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="2a739-1092">在适用于 iOS 的 Outlook 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="2a739-1092">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="2a739-p161">`getRegExMatches` 方法返回匹配在清单 XML 文件中的每个 `ItemHasRegularExpressionMatch` 或 `ItemHasKnownEntity` 规则元素中定义的正则表达式的字符串。对于 `ItemHasRegularExpressionMatch` 规则，匹配字符串必须发生在该规则指定的项目的属性中。`PropertyName` 简单类型定义支持的属性。</span><span class="sxs-lookup"><span data-stu-id="2a739-p161">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="2a739-1096">例如，考虑一个外接程序清单具有以下 `Rule` 元素：</span><span class="sxs-lookup"><span data-stu-id="2a739-1096">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="2a739-1097">从 `getRegExMatches` 返回的对象应有两个属性：`fruits` 和 `veggies`。</span><span class="sxs-lookup"><span data-stu-id="2a739-1097">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="2a739-p162">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。而是使用 [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) 方法检索整个正文。</span><span class="sxs-lookup"><span data-stu-id="2a739-p162">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="2a739-1101">Requirements</span><span class="sxs-lookup"><span data-stu-id="2a739-1101">Requirements</span></span>

|<span data-ttu-id="2a739-1102">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-1102">Requirement</span></span>|<span data-ttu-id="2a739-1103">值</span><span class="sxs-lookup"><span data-stu-id="2a739-1103">Value</span></span>|
|---|---|
|[<span data-ttu-id="2a739-1104">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2a739-1104">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2a739-1105">1.0</span><span class="sxs-lookup"><span data-stu-id="2a739-1105">1.0</span></span>|
|[<span data-ttu-id="2a739-1106">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2a739-1106">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2a739-1107">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2a739-1107">ReadItem</span></span>|
|[<span data-ttu-id="2a739-1108">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2a739-1108">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2a739-1109">阅读</span><span class="sxs-lookup"><span data-stu-id="2a739-1109">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="2a739-1110">返回：</span><span class="sxs-lookup"><span data-stu-id="2a739-1110">Returns:</span></span>

<span data-ttu-id="2a739-p163">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串数组的对象。每个数组的名称等于匹配 `ItemHasRegularExpressionMatch` 规则的 `RegExName` 属性或匹配 `ItemHasKnownEntity` 规则的 `FilterName` 属性的相应值。</span><span class="sxs-lookup"><span data-stu-id="2a739-p163">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="2a739-1113">

<dt>
类型</dt>


</span><span class="sxs-lookup"><span data-stu-id="2a739-1113">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="2a739-1114">对象</span><span class="sxs-lookup"><span data-stu-id="2a739-1114">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="2a739-1115">示例</span><span class="sxs-lookup"><span data-stu-id="2a739-1115">Example</span></span>

<span data-ttu-id="2a739-1116">以下示例显示了如何访问正则表达式规则元素 `fruits` 和 `veggies` 的匹配项的数组，这些元素在清单中指定。</span><span class="sxs-lookup"><span data-stu-id="2a739-1116">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="2a739-1117">getregexmatchesbyname （name） → (nullable) {数组。 < 字符串 >}</span><span class="sxs-lookup"><span data-stu-id="2a739-1117">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="2a739-1118">返回匹配在清单 XML 文件中定义的命名正则表达式的所选项目中的字符串值。</span><span class="sxs-lookup"><span data-stu-id="2a739-1118">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="2a739-1119">在适用于 iOS 的 Outlook 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="2a739-1119">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="2a739-1120">`getRegExMatchesByName` 方法返回匹配在具有指定 `RegExName` 元素值的清单 XML 文件中的 `ItemHasRegularExpressionMatch` 规则元素中定义的正则表达式的字符串。</span><span class="sxs-lookup"><span data-stu-id="2a739-1120">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="2a739-p164">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。</span><span class="sxs-lookup"><span data-stu-id="2a739-p164">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2a739-1123">参数：</span><span class="sxs-lookup"><span data-stu-id="2a739-1123">Parameters:</span></span>

|<span data-ttu-id="2a739-1124">名称</span><span class="sxs-lookup"><span data-stu-id="2a739-1124">Name</span></span>|<span data-ttu-id="2a739-1125">类型</span><span class="sxs-lookup"><span data-stu-id="2a739-1125">Type</span></span>|<span data-ttu-id="2a739-1126">说明</span><span class="sxs-lookup"><span data-stu-id="2a739-1126">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="2a739-1127">字符串</span><span class="sxs-lookup"><span data-stu-id="2a739-1127">String</span></span>|<span data-ttu-id="2a739-1128">定义筛选器匹配的 `ItemHasRegularExpressionMatch` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="2a739-1128">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2a739-1129">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-1129">Requirements</span></span>

|<span data-ttu-id="2a739-1130">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-1130">Requirement</span></span>|<span data-ttu-id="2a739-1131">值</span><span class="sxs-lookup"><span data-stu-id="2a739-1131">Value</span></span>|
|---|---|
|[<span data-ttu-id="2a739-1132">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2a739-1132">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2a739-1133">1.0</span><span class="sxs-lookup"><span data-stu-id="2a739-1133">1.0</span></span>|
|[<span data-ttu-id="2a739-1134">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2a739-1134">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2a739-1135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2a739-1135">ReadItem</span></span>|
|[<span data-ttu-id="2a739-1136">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2a739-1136">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2a739-1137">阅读</span><span class="sxs-lookup"><span data-stu-id="2a739-1137">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="2a739-1138">返回：</span><span class="sxs-lookup"><span data-stu-id="2a739-1138">Returns:</span></span>

<span data-ttu-id="2a739-1139">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串的数组。</span><span class="sxs-lookup"><span data-stu-id="2a739-1139">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="2a739-1140">

<dt>类型</dt>

</span><span class="sxs-lookup"><span data-stu-id="2a739-1140">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="2a739-1141">数组。 < 字符串 ></span><span class="sxs-lookup"><span data-stu-id="2a739-1141">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="2a739-1142">示例</span><span class="sxs-lookup"><span data-stu-id="2a739-1142">Example</span></span>

```
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="2a739-1143">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="2a739-1143">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="2a739-1144">以异步方式返回邮件的主题或正文中选定的数据。</span><span class="sxs-lookup"><span data-stu-id="2a739-1144">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="2a739-p165">如果没有选定内容，但光标位于正文或主题中，此方法将会为所选数据返回 null。如果选定的是字段，而不是正文或主题，则此方法返回 `InvalidSelection` 错误。</span><span class="sxs-lookup"><span data-stu-id="2a739-p165">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2a739-1147">参数：</span><span class="sxs-lookup"><span data-stu-id="2a739-1147">Parameters:</span></span>

|<span data-ttu-id="2a739-1148">名称</span><span class="sxs-lookup"><span data-stu-id="2a739-1148">Name</span></span>|<span data-ttu-id="2a739-1149">类型</span><span class="sxs-lookup"><span data-stu-id="2a739-1149">Type</span></span>|<span data-ttu-id="2a739-1150">属性</span><span class="sxs-lookup"><span data-stu-id="2a739-1150">Attributes</span></span>|<span data-ttu-id="2a739-1151">说明</span><span class="sxs-lookup"><span data-stu-id="2a739-1151">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="2a739-1152">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="2a739-1152">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="2a739-p166">请求数据的格式。如果为文本，则此方法返回纯文本作为字符串，删除任何显示的 HTML 标记。如果为 HTML，则此方法返回所选文本，不论是纯文本还是 HTML。</span><span class="sxs-lookup"><span data-stu-id="2a739-p166">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="2a739-1156">对象</span><span class="sxs-lookup"><span data-stu-id="2a739-1156">Object</span></span>|<span data-ttu-id="2a739-1157">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2a739-1157">&lt;optional&gt;</span></span>|<span data-ttu-id="2a739-1158">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="2a739-1158">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="2a739-1159">对象</span><span class="sxs-lookup"><span data-stu-id="2a739-1159">Object</span></span>|<span data-ttu-id="2a739-1160">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2a739-1160">&lt;optional&gt;</span></span>|<span data-ttu-id="2a739-1161">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="2a739-1161">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="2a739-1162">函数</span><span class="sxs-lookup"><span data-stu-id="2a739-1162">function</span></span>||<span data-ttu-id="2a739-1163">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="2a739-1163">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="2a739-1164">若要从回调方法访问所选数据，请调用 `asyncResult.value.data`。</span><span class="sxs-lookup"><span data-stu-id="2a739-1164">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="2a739-1165">若要访问所选内容来自源属性，请调用`asyncResult.value.sourceProperty`，其中将为`body`或`subject`。</span><span class="sxs-lookup"><span data-stu-id="2a739-1165">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2a739-1166">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-1166">Requirements</span></span>

|<span data-ttu-id="2a739-1167">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-1167">Requirement</span></span>|<span data-ttu-id="2a739-1168">值</span><span class="sxs-lookup"><span data-stu-id="2a739-1168">Value</span></span>|
|---|---|
|[<span data-ttu-id="2a739-1169">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2a739-1169">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2a739-1170">1.2</span><span class="sxs-lookup"><span data-stu-id="2a739-1170">1.2</span></span>|
|[<span data-ttu-id="2a739-1171">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2a739-1171">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2a739-1172">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="2a739-1172">ReadWriteItem</span></span>|
|[<span data-ttu-id="2a739-1173">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2a739-1173">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2a739-1174">撰写</span><span class="sxs-lookup"><span data-stu-id="2a739-1174">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="2a739-1175">返回：</span><span class="sxs-lookup"><span data-stu-id="2a739-1175">Returns:</span></span>

<span data-ttu-id="2a739-1176">作为字符串的所选数据的格式由 `coercionType` 确定。</span><span class="sxs-lookup"><span data-stu-id="2a739-1176">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="2a739-1177">

<dt>
类型</dt>


</span><span class="sxs-lookup"><span data-stu-id="2a739-1177">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="2a739-1178">字符串</span><span class="sxs-lookup"><span data-stu-id="2a739-1178">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="2a739-1179">示例</span><span class="sxs-lookup"><span data-stu-id="2a739-1179">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="2a739-1180">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="2a739-1180">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="2a739-p168">获取在用户已选择的突出显示匹配项中找到的实体。突出显示匹配项适用于[上下文外接程序](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins)。</span><span class="sxs-lookup"><span data-stu-id="2a739-p168">Gets the entities found in a highlighted match a user has selected. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="2a739-1183">在适用于 iOS 的 Outlook 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="2a739-1183">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="2a739-1184">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-1184">Requirements</span></span>

|<span data-ttu-id="2a739-1185">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-1185">Requirement</span></span>|<span data-ttu-id="2a739-1186">值</span><span class="sxs-lookup"><span data-stu-id="2a739-1186">Value</span></span>|
|---|---|
|[<span data-ttu-id="2a739-1187">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2a739-1187">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2a739-1188">1.6</span><span class="sxs-lookup"><span data-stu-id="2a739-1188">1.6</span></span>|
|[<span data-ttu-id="2a739-1189">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2a739-1189">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2a739-1190">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2a739-1190">ReadItem</span></span>|
|[<span data-ttu-id="2a739-1191">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2a739-1191">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2a739-1192">阅读</span><span class="sxs-lookup"><span data-stu-id="2a739-1192">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="2a739-1193">返回：</span><span class="sxs-lookup"><span data-stu-id="2a739-1193">Returns:</span></span>

<span data-ttu-id="2a739-1194">类型：[Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="2a739-1194">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="2a739-1195">示例</span><span class="sxs-lookup"><span data-stu-id="2a739-1195">Example</span></span>

<span data-ttu-id="2a739-1196">以下示例访问用户选择的突出显示匹配项中的地址实体。</span><span class="sxs-lookup"><span data-stu-id="2a739-1196">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="2a739-1197">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="2a739-1197">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="2a739-p169">返回突出显示匹配项中匹配在清单 XML 文件中定义的正则表达式的字符串值。突出显示匹配项适用于[上下文外接程序](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins)。</span><span class="sxs-lookup"><span data-stu-id="2a739-p169">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="2a739-1200">在适用于 iOS 的 Outlook 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="2a739-1200">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="2a739-p170">`getSelectedRegExMatches` 方法返回匹配在清单 XML 文件中的每个 `ItemHasRegularExpressionMatch` 或 `ItemHasKnownEntity` 规则元素中定义的正则表达式的字符串。对于 `ItemHasRegularExpressionMatch` 规则，匹配字符串必须发生在该规则指定的项目的属性中。`PropertyName` 简单类型定义支持的属性。</span><span class="sxs-lookup"><span data-stu-id="2a739-p170">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="2a739-1204">例如，考虑一个外接程序清单具有以下 `Rule` 元素：</span><span class="sxs-lookup"><span data-stu-id="2a739-1204">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="2a739-1205">从 `getRegExMatches` 返回的对象应有两个属性：`fruits` 和 `veggies`。</span><span class="sxs-lookup"><span data-stu-id="2a739-1205">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="2a739-p171">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。而是使用 [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) 方法检索整个正文。</span><span class="sxs-lookup"><span data-stu-id="2a739-p171">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="2a739-1209">Requirements</span><span class="sxs-lookup"><span data-stu-id="2a739-1209">Requirements</span></span>

|<span data-ttu-id="2a739-1210">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-1210">Requirement</span></span>|<span data-ttu-id="2a739-1211">值</span><span class="sxs-lookup"><span data-stu-id="2a739-1211">Value</span></span>|
|---|---|
|[<span data-ttu-id="2a739-1212">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2a739-1212">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2a739-1213">1.6</span><span class="sxs-lookup"><span data-stu-id="2a739-1213">1.6</span></span>|
|[<span data-ttu-id="2a739-1214">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2a739-1214">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2a739-1215">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2a739-1215">ReadItem</span></span>|
|[<span data-ttu-id="2a739-1216">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2a739-1216">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2a739-1217">阅读</span><span class="sxs-lookup"><span data-stu-id="2a739-1217">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="2a739-1218">返回：</span><span class="sxs-lookup"><span data-stu-id="2a739-1218">Returns:</span></span>

<span data-ttu-id="2a739-p172">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串数组的对象。每个数组的名称等于匹配 `ItemHasRegularExpressionMatch` 规则的 `RegExName` 属性或匹配 `ItemHasKnownEntity` 规则的 `FilterName` 属性的相应值。</span><span class="sxs-lookup"><span data-stu-id="2a739-p172">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="2a739-1221">示例</span><span class="sxs-lookup"><span data-stu-id="2a739-1221">Example</span></span>

<span data-ttu-id="2a739-1222">以下示例显示了如何访问正则表达式规则元素 `fruits` 和 `veggies` 的匹配项的数组，这些元素在清单中指定。</span><span class="sxs-lookup"><span data-stu-id="2a739-1222">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

#### <a name="getsharedpropertiesasyncoptions-callback"></a><span data-ttu-id="2a739-1223">getSharedPropertiesAsync ([选项] 回调)</span><span class="sxs-lookup"><span data-stu-id="2a739-1223">getSharedPropertiesAsync([options], callback)</span></span>

<span data-ttu-id="2a739-1224">获取共享的文件夹、 日历或邮箱中的所选的约会或邮件的属性。</span><span class="sxs-lookup"><span data-stu-id="2a739-1224">Gets the properties of the selected appointment or message in a shared folder, calendar, or mailbox.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2a739-1225">参数：</span><span class="sxs-lookup"><span data-stu-id="2a739-1225">Parameters:</span></span>

|<span data-ttu-id="2a739-1226">名称</span><span class="sxs-lookup"><span data-stu-id="2a739-1226">Name</span></span>|<span data-ttu-id="2a739-1227">类型</span><span class="sxs-lookup"><span data-stu-id="2a739-1227">Type</span></span>|<span data-ttu-id="2a739-1228">属性</span><span class="sxs-lookup"><span data-stu-id="2a739-1228">Attributes</span></span>|<span data-ttu-id="2a739-1229">说明</span><span class="sxs-lookup"><span data-stu-id="2a739-1229">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="2a739-1230">对象</span><span class="sxs-lookup"><span data-stu-id="2a739-1230">Object</span></span>|<span data-ttu-id="2a739-1231">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2a739-1231">&lt;optional&gt;</span></span>|<span data-ttu-id="2a739-1232">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="2a739-1232">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="2a739-1233">对象</span><span class="sxs-lookup"><span data-stu-id="2a739-1233">Object</span></span>|<span data-ttu-id="2a739-1234">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2a739-1234">&lt;optional&gt;</span></span>|<span data-ttu-id="2a739-1235">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="2a739-1235">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="2a739-1236">函数</span><span class="sxs-lookup"><span data-stu-id="2a739-1236">function</span></span>||<span data-ttu-id="2a739-1237">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="2a739-1237">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="2a739-1238">共享的属性提供为[`SharedProperties`](/javascript/api/outlook/office.sharedproperties)对象在`asyncResult.value`属性。</span><span class="sxs-lookup"><span data-stu-id="2a739-1238">The shared properties are provided as a [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="2a739-1239">此对象可用于获取项目的共享的属性。</span><span class="sxs-lookup"><span data-stu-id="2a739-1239">This object can be used to get the item's shared properties.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2a739-1240">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-1240">Requirements</span></span>

|<span data-ttu-id="2a739-1241">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-1241">Requirement</span></span>|<span data-ttu-id="2a739-1242">值</span><span class="sxs-lookup"><span data-stu-id="2a739-1242">Value</span></span>|
|---|---|
|[<span data-ttu-id="2a739-1243">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2a739-1243">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2a739-1244">预览</span><span class="sxs-lookup"><span data-stu-id="2a739-1244">Preview</span></span>|
|[<span data-ttu-id="2a739-1245">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2a739-1245">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2a739-1246">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2a739-1246">ReadItem</span></span>|
|[<span data-ttu-id="2a739-1247">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2a739-1247">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2a739-1248">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="2a739-1248">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="2a739-1249">示例</span><span class="sxs-lookup"><span data-stu-id="2a739-1249">Example</span></span>

```js
Office.context.mailbox.item.getSharedPropertiesAsync(callback);
function callback (asyncResult) {
  var context=asyncResult.context;
  var sharedProperties = asyncResult.value;
}
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="2a739-1250">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="2a739-1250">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="2a739-1251">异步加载所选项目上此外接程序的自定义属性。</span><span class="sxs-lookup"><span data-stu-id="2a739-1251">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="2a739-p174">自定义属性基于每个应用、每个项目存储为键/值对。此方法在回调中返回 `CustomProperties` 对象，该回调提供访问特定于当前项目和当前外接程序的自定义属性的方法。自定义属性未在项目上加密，因此这不应用作安全存储。</span><span class="sxs-lookup"><span data-stu-id="2a739-p174">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2a739-1255">参数：</span><span class="sxs-lookup"><span data-stu-id="2a739-1255">Parameters:</span></span>

|<span data-ttu-id="2a739-1256">名称</span><span class="sxs-lookup"><span data-stu-id="2a739-1256">Name</span></span>|<span data-ttu-id="2a739-1257">类型</span><span class="sxs-lookup"><span data-stu-id="2a739-1257">Type</span></span>|<span data-ttu-id="2a739-1258">属性</span><span class="sxs-lookup"><span data-stu-id="2a739-1258">Attributes</span></span>|<span data-ttu-id="2a739-1259">说明</span><span class="sxs-lookup"><span data-stu-id="2a739-1259">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="2a739-1260">函数</span><span class="sxs-lookup"><span data-stu-id="2a739-1260">function</span></span>||<span data-ttu-id="2a739-1261">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="2a739-1261">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="2a739-1262">自定义属性作为 `asyncResult.value` 属性中的 [`CustomProperties`](/javascript/api/outlook/office.customproperties) 对象提供。</span><span class="sxs-lookup"><span data-stu-id="2a739-1262">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="2a739-1263">此对象可用于获取、 设置，和从项目中删除自定义属性并将更改保存到自定义属性设置回服务器。</span><span class="sxs-lookup"><span data-stu-id="2a739-1263">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="2a739-1264">对象</span><span class="sxs-lookup"><span data-stu-id="2a739-1264">Object</span></span>|<span data-ttu-id="2a739-1265">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2a739-1265">&lt;optional&gt;</span></span>|<span data-ttu-id="2a739-1266">开发人员可以提供他们想要访问的回调函数中的任何对象。</span><span class="sxs-lookup"><span data-stu-id="2a739-1266">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="2a739-1267">此对象可以访问`asyncResult.asyncContext`的回调函数中的属性。</span><span class="sxs-lookup"><span data-stu-id="2a739-1267">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2a739-1268">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-1268">Requirements</span></span>

|<span data-ttu-id="2a739-1269">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-1269">Requirement</span></span>|<span data-ttu-id="2a739-1270">值</span><span class="sxs-lookup"><span data-stu-id="2a739-1270">Value</span></span>|
|---|---|
|[<span data-ttu-id="2a739-1271">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2a739-1271">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2a739-1272">1.0</span><span class="sxs-lookup"><span data-stu-id="2a739-1272">1.0</span></span>|
|[<span data-ttu-id="2a739-1273">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2a739-1273">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2a739-1274">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2a739-1274">ReadItem</span></span>|
|[<span data-ttu-id="2a739-1275">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2a739-1275">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2a739-1276">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="2a739-1276">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="2a739-1277">示例</span><span class="sxs-lookup"><span data-stu-id="2a739-1277">Example</span></span>

<span data-ttu-id="2a739-p177">以下代码示例显示了如何使用 `loadCustomPropertiesAsync` 方法异步加载特定于当前项目的自定义属性。该示例还显示了如何使用 `CustomProperties.saveAsync` 方法将这些属性重新保存到服务器。加载自定义属性后，该代码示例将使用 `CustomProperties.get` 方法读取自定义属性 `myProp`，使用 `CustomProperties.set` 方法写入自定义属性 `otherProp`，最后调用 `saveAsync` 方法保存这些自定义属性。</span><span class="sxs-lookup"><span data-stu-id="2a739-p177">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="2a739-1281">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="2a739-1281">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="2a739-1282">将附件从邮件或约会中删除。</span><span class="sxs-lookup"><span data-stu-id="2a739-1282">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="2a739-p178">`removeAttachmentAsync` 方法删除项目中带指定标识符的附件。最佳做法是，仅当同一个邮件应用程序在同一会话中添加了一个附件时，你才应使用该附件标识符来删除该附件。在 Outlook Web App 和适用于设备的 OWA 中，附件标识符只在同一个会话中才有效。当用户关闭应用程序，或者如果用户开始在内嵌窗体中撰写，并在随后弹出的内嵌窗体中继续在单独的窗口撰写时，会话即结束。</span><span class="sxs-lookup"><span data-stu-id="2a739-p178">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2a739-1287">参数：</span><span class="sxs-lookup"><span data-stu-id="2a739-1287">Parameters:</span></span>

|<span data-ttu-id="2a739-1288">名称</span><span class="sxs-lookup"><span data-stu-id="2a739-1288">Name</span></span>|<span data-ttu-id="2a739-1289">类型</span><span class="sxs-lookup"><span data-stu-id="2a739-1289">Type</span></span>|<span data-ttu-id="2a739-1290">属性</span><span class="sxs-lookup"><span data-stu-id="2a739-1290">Attributes</span></span>|<span data-ttu-id="2a739-1291">说明</span><span class="sxs-lookup"><span data-stu-id="2a739-1291">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="2a739-1292">字符串</span><span class="sxs-lookup"><span data-stu-id="2a739-1292">String</span></span>||<span data-ttu-id="2a739-p179">要删除的附件的标识符。字符串的最大长度为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="2a739-p179">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`|<span data-ttu-id="2a739-1295">对象</span><span class="sxs-lookup"><span data-stu-id="2a739-1295">Object</span></span>|<span data-ttu-id="2a739-1296">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2a739-1296">&lt;optional&gt;</span></span>|<span data-ttu-id="2a739-1297">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="2a739-1297">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="2a739-1298">对象</span><span class="sxs-lookup"><span data-stu-id="2a739-1298">Object</span></span>|<span data-ttu-id="2a739-1299">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2a739-1299">&lt;optional&gt;</span></span>|<span data-ttu-id="2a739-1300">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="2a739-1300">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="2a739-1301">函数</span><span class="sxs-lookup"><span data-stu-id="2a739-1301">function</span></span>|<span data-ttu-id="2a739-1302">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2a739-1302">&lt;optional&gt;</span></span>|<span data-ttu-id="2a739-1303">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="2a739-1303">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="2a739-1304">如果删除附件失败，`asyncResult.error` 属性将包含一个说明失败原因的错误代码。</span><span class="sxs-lookup"><span data-stu-id="2a739-1304">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="2a739-1305">错误</span><span class="sxs-lookup"><span data-stu-id="2a739-1305">Errors</span></span>

|<span data-ttu-id="2a739-1306">错误代码</span><span class="sxs-lookup"><span data-stu-id="2a739-1306">Error code</span></span>|<span data-ttu-id="2a739-1307">说明</span><span class="sxs-lookup"><span data-stu-id="2a739-1307">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="2a739-1308">附件标识符不存在。</span><span class="sxs-lookup"><span data-stu-id="2a739-1308">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2a739-1309">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-1309">Requirements</span></span>

|<span data-ttu-id="2a739-1310">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-1310">Requirement</span></span>|<span data-ttu-id="2a739-1311">值</span><span class="sxs-lookup"><span data-stu-id="2a739-1311">Value</span></span>|
|---|---|
|[<span data-ttu-id="2a739-1312">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2a739-1312">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2a739-1313">1.1</span><span class="sxs-lookup"><span data-stu-id="2a739-1313">1.1</span></span>|
|[<span data-ttu-id="2a739-1314">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2a739-1314">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2a739-1315">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="2a739-1315">ReadWriteItem</span></span>|
|[<span data-ttu-id="2a739-1316">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2a739-1316">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2a739-1317">撰写</span><span class="sxs-lookup"><span data-stu-id="2a739-1317">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="2a739-1318">示例</span><span class="sxs-lookup"><span data-stu-id="2a739-1318">Example</span></span>

<span data-ttu-id="2a739-1319">以下代码删除包含标识符 '0' 的附件。</span><span class="sxs-lookup"><span data-stu-id="2a739-1319">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="removehandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="2a739-1320">removeHandlerAsync (eventType，处理程序中，[选项] [回调])</span><span class="sxs-lookup"><span data-stu-id="2a739-1320">removeHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="2a739-1321">删除事件处理程序支持的事件。</span><span class="sxs-lookup"><span data-stu-id="2a739-1321">Removes an event handler for a supported event.</span></span>

<span data-ttu-id="2a739-1322">当前受支持的事件类型的`Office.EventType.AppointmentTimeChanged`， `Office.EventType.RecipientsChanged`，和`Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="2a739-1322">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="2a739-1323">参数：</span><span class="sxs-lookup"><span data-stu-id="2a739-1323">Parameters:</span></span>

| <span data-ttu-id="2a739-1324">名称</span><span class="sxs-lookup"><span data-stu-id="2a739-1324">Name</span></span> | <span data-ttu-id="2a739-1325">类型</span><span class="sxs-lookup"><span data-stu-id="2a739-1325">Type</span></span> | <span data-ttu-id="2a739-1326">属性</span><span class="sxs-lookup"><span data-stu-id="2a739-1326">Attributes</span></span> | <span data-ttu-id="2a739-1327">说明</span><span class="sxs-lookup"><span data-stu-id="2a739-1327">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="2a739-1328">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="2a739-1328">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="2a739-1329">应调用处理程序的事件。</span><span class="sxs-lookup"><span data-stu-id="2a739-1329">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="2a739-1330">函数</span><span class="sxs-lookup"><span data-stu-id="2a739-1330">Function</span></span> || <span data-ttu-id="2a739-p180">用于处理事件的函数。此函数必须接受一个参数，即对象文本。参数上的 `type` 属性将匹配传递给 `removeHandlerAsync` 的 `eventType` 参数。</span><span class="sxs-lookup"><span data-stu-id="2a739-p180">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `removeHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="2a739-1334">Object</span><span class="sxs-lookup"><span data-stu-id="2a739-1334">Object</span></span> | <span data-ttu-id="2a739-1335">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2a739-1335">&lt;optional&gt;</span></span> | <span data-ttu-id="2a739-1336">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="2a739-1336">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="2a739-1337">对象</span><span class="sxs-lookup"><span data-stu-id="2a739-1337">Object</span></span> | <span data-ttu-id="2a739-1338">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2a739-1338">&lt;optional&gt;</span></span> | <span data-ttu-id="2a739-1339">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="2a739-1339">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="2a739-1340">函数</span><span class="sxs-lookup"><span data-stu-id="2a739-1340">function</span></span>| <span data-ttu-id="2a739-1341">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2a739-1341">&lt;optional&gt;</span></span>|<span data-ttu-id="2a739-1342">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="2a739-1342">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2a739-1343">Requirements</span><span class="sxs-lookup"><span data-stu-id="2a739-1343">Requirements</span></span>

|<span data-ttu-id="2a739-1344">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-1344">Requirement</span></span>| <span data-ttu-id="2a739-1345">值</span><span class="sxs-lookup"><span data-stu-id="2a739-1345">Value</span></span>|
|---|---|
|[<span data-ttu-id="2a739-1346">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2a739-1346">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2a739-1347">1.7</span><span class="sxs-lookup"><span data-stu-id="2a739-1347">1.7</span></span> |
|[<span data-ttu-id="2a739-1348">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2a739-1348">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2a739-1349">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2a739-1349">ReadItem</span></span> |
|[<span data-ttu-id="2a739-1350">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2a739-1350">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2a739-1351">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="2a739-1351">Compose or read</span></span> |

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="2a739-1352">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="2a739-1352">saveAsync([options], callback)</span></span>

<span data-ttu-id="2a739-1353">异步保存项目。</span><span class="sxs-lookup"><span data-stu-id="2a739-1353">Asynchronously saves an item.</span></span>

<span data-ttu-id="2a739-p181">调用时，此方法将当前邮件保存为草稿，并通过回调方法返回项目 ID。在 Outlook Web App 或 Outlook 联机模式下，该项目被保存到服务器中。在 Outlook 缓存模式下，该项目被保存到本地缓存中。</span><span class="sxs-lookup"><span data-stu-id="2a739-p181">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="2a739-1357">如果加载项调用`saveAsync`中的项目在撰写模式下才能获取`itemId`若要使用 EWS 或 REST API，请注意，缓存模式 Outlook 时，可能需要一些时间才能项目实际同步到服务器。</span><span class="sxs-lookup"><span data-stu-id="2a739-1357">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="2a739-1358">直到该项目同步，使用`itemId`将返回错误。</span><span class="sxs-lookup"><span data-stu-id="2a739-1358">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="2a739-p183">由于约会没有草稿状态，如果以撰写模式在约会中调用 `saveAsync`，则该项将被保存为用户日历中的正常约会。对于之前未保存过的新约会，则不会发送邀请。保存现有约会将向添加或删除的与会者发送更新。</span><span class="sxs-lookup"><span data-stu-id="2a739-p183">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="2a739-1362">以下客户端具有不同行为的`saveAsync`上约会的撰写模式下：</span><span class="sxs-lookup"><span data-stu-id="2a739-1362">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="2a739-1363">Mac Outlook 不支持`saveAsync`中的会议在撰写模式。</span><span class="sxs-lookup"><span data-stu-id="2a739-1363">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="2a739-1364">调用`saveAsync`上 Mac Outlook 中的会议，则将返回错误。</span><span class="sxs-lookup"><span data-stu-id="2a739-1364">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="2a739-1365">在 web 上的 outlook 始终发送的邀请或更新何时`saveAsync`上约会调用在撰写模式。</span><span class="sxs-lookup"><span data-stu-id="2a739-1365">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2a739-1366">参数：</span><span class="sxs-lookup"><span data-stu-id="2a739-1366">Parameters:</span></span>

|<span data-ttu-id="2a739-1367">名称</span><span class="sxs-lookup"><span data-stu-id="2a739-1367">Name</span></span>|<span data-ttu-id="2a739-1368">类型</span><span class="sxs-lookup"><span data-stu-id="2a739-1368">Type</span></span>|<span data-ttu-id="2a739-1369">属性</span><span class="sxs-lookup"><span data-stu-id="2a739-1369">Attributes</span></span>|<span data-ttu-id="2a739-1370">说明</span><span class="sxs-lookup"><span data-stu-id="2a739-1370">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="2a739-1371">对象</span><span class="sxs-lookup"><span data-stu-id="2a739-1371">Object</span></span>|<span data-ttu-id="2a739-1372">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2a739-1372">&lt;optional&gt;</span></span>|<span data-ttu-id="2a739-1373">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="2a739-1373">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="2a739-1374">对象</span><span class="sxs-lookup"><span data-stu-id="2a739-1374">Object</span></span>|<span data-ttu-id="2a739-1375">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2a739-1375">&lt;optional&gt;</span></span>|<span data-ttu-id="2a739-1376">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="2a739-1376">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="2a739-1377">函数</span><span class="sxs-lookup"><span data-stu-id="2a739-1377">function</span></span>||<span data-ttu-id="2a739-1378">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="2a739-1378">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="2a739-1379">中提供的项标识符是成功，`asyncResult.value`属性。</span><span class="sxs-lookup"><span data-stu-id="2a739-1379">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2a739-1380">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-1380">Requirements</span></span>

|<span data-ttu-id="2a739-1381">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-1381">Requirement</span></span>|<span data-ttu-id="2a739-1382">值</span><span class="sxs-lookup"><span data-stu-id="2a739-1382">Value</span></span>|
|---|---|
|[<span data-ttu-id="2a739-1383">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2a739-1383">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2a739-1384">1.3</span><span class="sxs-lookup"><span data-stu-id="2a739-1384">1.3</span></span>|
|[<span data-ttu-id="2a739-1385">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2a739-1385">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2a739-1386">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="2a739-1386">ReadWriteItem</span></span>|
|[<span data-ttu-id="2a739-1387">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2a739-1387">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2a739-1388">撰写</span><span class="sxs-lookup"><span data-stu-id="2a739-1388">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="2a739-1389">示例</span><span class="sxs-lookup"><span data-stu-id="2a739-1389">Examples</span></span>

```
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="2a739-p185">下面是传递给回调函数的 `result` 参数的示例。`value` 属性包含的项目的项目 ID。</span><span class="sxs-lookup"><span data-stu-id="2a739-p185">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="2a739-1392">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="2a739-1392">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="2a739-1393">以异步方式将数据插入到邮件的正文或主题中。</span><span class="sxs-lookup"><span data-stu-id="2a739-1393">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="2a739-p186">`setSelectedDataAsync` 方法将指定字符串插入到项目主题或正文的光标位置，或者，如果在编辑器中已选择文本，则该方法将替换选择的文本。如果光标不在正文或主题字段中，则返回错误。插入之后，光标会位于插入内容的末尾。</span><span class="sxs-lookup"><span data-stu-id="2a739-p186">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2a739-1397">参数：</span><span class="sxs-lookup"><span data-stu-id="2a739-1397">Parameters:</span></span>

|<span data-ttu-id="2a739-1398">名称</span><span class="sxs-lookup"><span data-stu-id="2a739-1398">Name</span></span>|<span data-ttu-id="2a739-1399">类型</span><span class="sxs-lookup"><span data-stu-id="2a739-1399">Type</span></span>|<span data-ttu-id="2a739-1400">属性</span><span class="sxs-lookup"><span data-stu-id="2a739-1400">Attributes</span></span>|<span data-ttu-id="2a739-1401">说明</span><span class="sxs-lookup"><span data-stu-id="2a739-1401">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="2a739-1402">字符串</span><span class="sxs-lookup"><span data-stu-id="2a739-1402">String</span></span>||<span data-ttu-id="2a739-p187">要插入的数据。数据不得超过 1,000,000 个字符。如果传入的数据超过 1,000,000 个字符，则会引发 `ArgumentOutOfRange` 异常。</span><span class="sxs-lookup"><span data-stu-id="2a739-p187">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="2a739-1406">Object</span><span class="sxs-lookup"><span data-stu-id="2a739-1406">Object</span></span>|<span data-ttu-id="2a739-1407">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2a739-1407">&lt;optional&gt;</span></span>|<span data-ttu-id="2a739-1408">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="2a739-1408">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="2a739-1409">对象</span><span class="sxs-lookup"><span data-stu-id="2a739-1409">Object</span></span>|<span data-ttu-id="2a739-1410">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2a739-1410">&lt;optional&gt;</span></span>|<span data-ttu-id="2a739-1411">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="2a739-1411">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="2a739-1412">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="2a739-1412">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="2a739-1413">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="2a739-1413">&lt;optional&gt;</span></span>|<span data-ttu-id="2a739-p188">如果为 `text`，则在 Outlook Web App 和 Outlook 中应用当前样式。如果该字段是 HTML 编辑器，则仅插入文本数据，即使数据为 HTML。</span><span class="sxs-lookup"><span data-stu-id="2a739-p188">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="2a739-p189">如果 `html` 和该字段支持 HTML（主题不支持），则在 Outlook Web App 中应用当前样式，而在 Outlook 中应用默认样式。如果该字段是文本字段，则返回 `InvalidDataFormat` 错误。</span><span class="sxs-lookup"><span data-stu-id="2a739-p189">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="2a739-1418">如果未设置 `coercionType`，则结果取决于该字段：如果该字段是 HTML，则使用 HTML；如果该字段是文本，则使用纯文本。</span><span class="sxs-lookup"><span data-stu-id="2a739-1418">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="2a739-1419">function</span><span class="sxs-lookup"><span data-stu-id="2a739-1419">function</span></span>||<span data-ttu-id="2a739-1420">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="2a739-1420">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2a739-1421">Requirements</span><span class="sxs-lookup"><span data-stu-id="2a739-1421">Requirements</span></span>

|<span data-ttu-id="2a739-1422">要求</span><span class="sxs-lookup"><span data-stu-id="2a739-1422">Requirement</span></span>|<span data-ttu-id="2a739-1423">值</span><span class="sxs-lookup"><span data-stu-id="2a739-1423">Value</span></span>|
|---|---|
|[<span data-ttu-id="2a739-1424">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2a739-1424">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2a739-1425">1.2</span><span class="sxs-lookup"><span data-stu-id="2a739-1425">1.2</span></span>|
|[<span data-ttu-id="2a739-1426">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2a739-1426">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2a739-1427">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="2a739-1427">ReadWriteItem</span></span>|
|[<span data-ttu-id="2a739-1428">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2a739-1428">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="2a739-1429">撰写</span><span class="sxs-lookup"><span data-stu-id="2a739-1429">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="2a739-1430">示例</span><span class="sxs-lookup"><span data-stu-id="2a739-1430">Example</span></span>

```
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```