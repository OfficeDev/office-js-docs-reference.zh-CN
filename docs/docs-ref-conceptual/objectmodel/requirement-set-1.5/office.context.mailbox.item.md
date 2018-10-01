
# <a name="item"></a><span data-ttu-id="72c29-101">item</span><span class="sxs-lookup"><span data-stu-id="72c29-101">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="72c29-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="72c29-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="72c29-p101">`item` 命名空间用于访问当前选定的邮件、会议请求或约会。可以通过使用 [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype) 属性确定 `item` 的类型。</span><span class="sxs-lookup"><span data-stu-id="72c29-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="72c29-105">Requirements</span><span class="sxs-lookup"><span data-stu-id="72c29-105">Requirements</span></span>

|<span data-ttu-id="72c29-106">要求</span><span class="sxs-lookup"><span data-stu-id="72c29-106">Requirement</span></span>| <span data-ttu-id="72c29-107">值</span><span class="sxs-lookup"><span data-stu-id="72c29-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="72c29-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="72c29-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="72c29-109">1.0</span><span class="sxs-lookup"><span data-stu-id="72c29-109">1.0</span></span>|
|[<span data-ttu-id="72c29-110">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="72c29-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="72c29-111">受限</span><span class="sxs-lookup"><span data-stu-id="72c29-111">Restricted</span></span>|
|[<span data-ttu-id="72c29-112">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="72c29-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="72c29-113">Compose 或 Read</span><span class="sxs-lookup"><span data-stu-id="72c29-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="72c29-114">成员和方法</span><span class="sxs-lookup"><span data-stu-id="72c29-114">Members and methods</span></span>

| <span data-ttu-id="72c29-115">成员</span><span class="sxs-lookup"><span data-stu-id="72c29-115">Member</span></span> | <span data-ttu-id="72c29-116">类型</span><span class="sxs-lookup"><span data-stu-id="72c29-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="72c29-117">attachments</span><span class="sxs-lookup"><span data-stu-id="72c29-117">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlook15officeattachmentdetails) | <span data-ttu-id="72c29-118">成员</span><span class="sxs-lookup"><span data-stu-id="72c29-118">Member</span></span> |
| [<span data-ttu-id="72c29-119">bcc</span><span class="sxs-lookup"><span data-stu-id="72c29-119">bcc</span></span>](#bcc-recipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="72c29-120">成员</span><span class="sxs-lookup"><span data-stu-id="72c29-120">Member</span></span> |
| [<span data-ttu-id="72c29-121">body</span><span class="sxs-lookup"><span data-stu-id="72c29-121">body</span></span>](#body-bodyjavascriptapioutlook15officebody) | <span data-ttu-id="72c29-122">成员</span><span class="sxs-lookup"><span data-stu-id="72c29-122">Member</span></span> |
| [<span data-ttu-id="72c29-123">cc</span><span class="sxs-lookup"><span data-stu-id="72c29-123">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="72c29-124">成员</span><span class="sxs-lookup"><span data-stu-id="72c29-124">Member</span></span> |
| [<span data-ttu-id="72c29-125">conversationId</span><span class="sxs-lookup"><span data-stu-id="72c29-125">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="72c29-126">成员</span><span class="sxs-lookup"><span data-stu-id="72c29-126">Member</span></span> |
| [<span data-ttu-id="72c29-127">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="72c29-127">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="72c29-128">成员</span><span class="sxs-lookup"><span data-stu-id="72c29-128">Member</span></span> |
| [<span data-ttu-id="72c29-129">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="72c29-129">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="72c29-130">成员</span><span class="sxs-lookup"><span data-stu-id="72c29-130">Member</span></span> |
| [<span data-ttu-id="72c29-131">end</span><span class="sxs-lookup"><span data-stu-id="72c29-131">end</span></span>](#end-datetimejavascriptapioutlook15officetime) | <span data-ttu-id="72c29-132">成员</span><span class="sxs-lookup"><span data-stu-id="72c29-132">Member</span></span> |
| [<span data-ttu-id="72c29-133">from</span><span class="sxs-lookup"><span data-stu-id="72c29-133">from</span></span>](#from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) | <span data-ttu-id="72c29-134">成员</span><span class="sxs-lookup"><span data-stu-id="72c29-134">Member</span></span> |
| [<span data-ttu-id="72c29-135">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="72c29-135">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="72c29-136">成员</span><span class="sxs-lookup"><span data-stu-id="72c29-136">Member</span></span> |
| [<span data-ttu-id="72c29-137">itemClass</span><span class="sxs-lookup"><span data-stu-id="72c29-137">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="72c29-138">成员</span><span class="sxs-lookup"><span data-stu-id="72c29-138">Member</span></span> |
| [<span data-ttu-id="72c29-139">itemId</span><span class="sxs-lookup"><span data-stu-id="72c29-139">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="72c29-140">成员</span><span class="sxs-lookup"><span data-stu-id="72c29-140">Member</span></span> |
| [<span data-ttu-id="72c29-141">itemType</span><span class="sxs-lookup"><span data-stu-id="72c29-141">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype) | <span data-ttu-id="72c29-142">成员</span><span class="sxs-lookup"><span data-stu-id="72c29-142">Member</span></span> |
| [<span data-ttu-id="72c29-143">location</span><span class="sxs-lookup"><span data-stu-id="72c29-143">location</span></span>](#location-stringlocationjavascriptapioutlook15officelocation) | <span data-ttu-id="72c29-144">成员</span><span class="sxs-lookup"><span data-stu-id="72c29-144">Member</span></span> |
| [<span data-ttu-id="72c29-145">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="72c29-145">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="72c29-146">成员</span><span class="sxs-lookup"><span data-stu-id="72c29-146">Member</span></span> |
| [<span data-ttu-id="72c29-147">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="72c29-147">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlook15officenotificationmessages) | <span data-ttu-id="72c29-148">成员</span><span class="sxs-lookup"><span data-stu-id="72c29-148">Member</span></span> |
| [<span data-ttu-id="72c29-149">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="72c29-149">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="72c29-150">成员</span><span class="sxs-lookup"><span data-stu-id="72c29-150">Member</span></span> |
| [<span data-ttu-id="72c29-151">organizer</span><span class="sxs-lookup"><span data-stu-id="72c29-151">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) | <span data-ttu-id="72c29-152">成员</span><span class="sxs-lookup"><span data-stu-id="72c29-152">Member</span></span> |
| [<span data-ttu-id="72c29-153">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="72c29-153">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="72c29-154">成员</span><span class="sxs-lookup"><span data-stu-id="72c29-154">Member</span></span> |
| [<span data-ttu-id="72c29-155">sender</span><span class="sxs-lookup"><span data-stu-id="72c29-155">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) | <span data-ttu-id="72c29-156">成员</span><span class="sxs-lookup"><span data-stu-id="72c29-156">Member</span></span> |
| [<span data-ttu-id="72c29-157">start</span><span class="sxs-lookup"><span data-stu-id="72c29-157">start</span></span>](#start-datetimejavascriptapioutlook15officetime) | <span data-ttu-id="72c29-158">成员</span><span class="sxs-lookup"><span data-stu-id="72c29-158">Member</span></span> |
| [<span data-ttu-id="72c29-159">subject</span><span class="sxs-lookup"><span data-stu-id="72c29-159">subject</span></span>](#subject-stringsubjectjavascriptapioutlook15officesubject) | <span data-ttu-id="72c29-160">成员</span><span class="sxs-lookup"><span data-stu-id="72c29-160">Member</span></span> |
| [<span data-ttu-id="72c29-161">to</span><span class="sxs-lookup"><span data-stu-id="72c29-161">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="72c29-162">成员</span><span class="sxs-lookup"><span data-stu-id="72c29-162">Member</span></span> |
| [<span data-ttu-id="72c29-163">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="72c29-163">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="72c29-164">方法</span><span class="sxs-lookup"><span data-stu-id="72c29-164">Method</span></span> |
| [<span data-ttu-id="72c29-165">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="72c29-165">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="72c29-166">方法</span><span class="sxs-lookup"><span data-stu-id="72c29-166">Method</span></span> |
| [<span data-ttu-id="72c29-167">close</span><span class="sxs-lookup"><span data-stu-id="72c29-167">close</span></span>](#close) | <span data-ttu-id="72c29-168">方法</span><span class="sxs-lookup"><span data-stu-id="72c29-168">Method</span></span> |
| [<span data-ttu-id="72c29-169">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="72c29-169">displayReplyAllForm</span></span>](#displayreplyallformformdata) | <span data-ttu-id="72c29-170">方法</span><span class="sxs-lookup"><span data-stu-id="72c29-170">Method</span></span> |
| [<span data-ttu-id="72c29-171">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="72c29-171">displayReplyForm</span></span>](#displayreplyformformdata) | <span data-ttu-id="72c29-172">方法</span><span class="sxs-lookup"><span data-stu-id="72c29-172">Method</span></span> |
| [<span data-ttu-id="72c29-173">getEntities</span><span class="sxs-lookup"><span data-stu-id="72c29-173">getEntities</span></span>](#getentities--entitiesjavascriptapioutlook15officeentities) | <span data-ttu-id="72c29-174">方法</span><span class="sxs-lookup"><span data-stu-id="72c29-174">Method</span></span> |
| [<span data-ttu-id="72c29-175">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="72c29-175">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion) | <span data-ttu-id="72c29-176">方法</span><span class="sxs-lookup"><span data-stu-id="72c29-176">Method</span></span> |
| [<span data-ttu-id="72c29-177">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="72c29-177">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion) | <span data-ttu-id="72c29-178">方法</span><span class="sxs-lookup"><span data-stu-id="72c29-178">Method</span></span> |
| [<span data-ttu-id="72c29-179">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="72c29-179">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="72c29-180">方法</span><span class="sxs-lookup"><span data-stu-id="72c29-180">Method</span></span> |
| [<span data-ttu-id="72c29-181">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="72c29-181">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="72c29-182">方法</span><span class="sxs-lookup"><span data-stu-id="72c29-182">Method</span></span> |
| [<span data-ttu-id="72c29-183">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="72c29-183">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="72c29-184">方法</span><span class="sxs-lookup"><span data-stu-id="72c29-184">Method</span></span> |
| [<span data-ttu-id="72c29-185">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="72c29-185">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="72c29-186">方法</span><span class="sxs-lookup"><span data-stu-id="72c29-186">Method</span></span> |
| [<span data-ttu-id="72c29-187">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="72c29-187">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="72c29-188">方法</span><span class="sxs-lookup"><span data-stu-id="72c29-188">Method</span></span> |
| [<span data-ttu-id="72c29-189">saveAsync</span><span class="sxs-lookup"><span data-stu-id="72c29-189">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="72c29-190">方法</span><span class="sxs-lookup"><span data-stu-id="72c29-190">Method</span></span> |
| [<span data-ttu-id="72c29-191">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="72c29-191">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="72c29-192">方法</span><span class="sxs-lookup"><span data-stu-id="72c29-192">Method</span></span> |

### <a name="example"></a><span data-ttu-id="72c29-193">示例</span><span class="sxs-lookup"><span data-stu-id="72c29-193">Example</span></span>

<span data-ttu-id="72c29-194">以下 JavaScript 代码示例显示了如何访问 Outlook 中当前项目的 `subject` 属性。</span><span class="sxs-lookup"><span data-stu-id="72c29-194">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="72c29-195">成员</span><span class="sxs-lookup"><span data-stu-id="72c29-195">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook15officeattachmentdetails"></a><span data-ttu-id="72c29-196">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="72c29-196">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span></span>

<span data-ttu-id="72c29-p102">获取项目的附件数组。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="72c29-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="72c29-199">某些类型的文件被阻止 Outlook 由于潜在的安全问题，并因此不返回。</span><span class="sxs-lookup"><span data-stu-id="72c29-199">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="72c29-200">有关详细信息，请参阅[在 Outlook 中被阻止的附件](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)。</span><span class="sxs-lookup"><span data-stu-id="72c29-200">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="72c29-201">类型:</span><span class="sxs-lookup"><span data-stu-id="72c29-201">Type:</span></span>

*   <span data-ttu-id="72c29-202">Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="72c29-202">Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="72c29-203">要求</span><span class="sxs-lookup"><span data-stu-id="72c29-203">Requirements</span></span>

|<span data-ttu-id="72c29-204">要求</span><span class="sxs-lookup"><span data-stu-id="72c29-204">Requirement</span></span>| <span data-ttu-id="72c29-205">值</span><span class="sxs-lookup"><span data-stu-id="72c29-205">Value</span></span>|
|---|---|
|[<span data-ttu-id="72c29-206">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="72c29-206">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="72c29-207">1.0</span><span class="sxs-lookup"><span data-stu-id="72c29-207">1.0</span></span>|
|[<span data-ttu-id="72c29-208">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="72c29-208">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="72c29-209">ReadItem</span><span class="sxs-lookup"><span data-stu-id="72c29-209">ReadItem</span></span>|
|[<span data-ttu-id="72c29-210">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="72c29-210">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="72c29-211">阅读</span><span class="sxs-lookup"><span data-stu-id="72c29-211">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="72c29-212">示例</span><span class="sxs-lookup"><span data-stu-id="72c29-212">Example</span></span>

<span data-ttu-id="72c29-213">以下代码使用当前项目上所有附件的详细信息构成 HTML 字符串。</span><span class="sxs-lookup"><span data-stu-id="72c29-213">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="72c29-214">bcc :[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="72c29-214">bcc :[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="72c29-215">获取一个对象，提供用于获取或更新邮件的密件抄送 （密件抄送副本） 行上的收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="72c29-215">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="72c29-216">仅限撰写模式。</span><span class="sxs-lookup"><span data-stu-id="72c29-216">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="72c29-217">类型：</span><span class="sxs-lookup"><span data-stu-id="72c29-217">Type:</span></span>

*   [<span data-ttu-id="72c29-218">收件人</span><span class="sxs-lookup"><span data-stu-id="72c29-218">Recipients</span></span>](/javascript/api/outlook_1_5/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="72c29-219">要求</span><span class="sxs-lookup"><span data-stu-id="72c29-219">Requirements</span></span>

|<span data-ttu-id="72c29-220">要求</span><span class="sxs-lookup"><span data-stu-id="72c29-220">Requirement</span></span>| <span data-ttu-id="72c29-221">值</span><span class="sxs-lookup"><span data-stu-id="72c29-221">Value</span></span>|
|---|---|
|[<span data-ttu-id="72c29-222">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="72c29-222">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="72c29-223">1.1</span><span class="sxs-lookup"><span data-stu-id="72c29-223">1.1</span></span>|
|[<span data-ttu-id="72c29-224">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="72c29-224">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="72c29-225">ReadItem</span><span class="sxs-lookup"><span data-stu-id="72c29-225">ReadItem</span></span>|
|[<span data-ttu-id="72c29-226">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="72c29-226">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="72c29-227">撰写</span><span class="sxs-lookup"><span data-stu-id="72c29-227">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="72c29-228">示例</span><span class="sxs-lookup"><span data-stu-id="72c29-228">Example</span></span>

```
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook15officebody"></a><span data-ttu-id="72c29-229">body :[Body](/javascript/api/outlook_1_5/office.body)</span><span class="sxs-lookup"><span data-stu-id="72c29-229">body :[Body](/javascript/api/outlook_1_5/office.body)</span></span>

<span data-ttu-id="72c29-230">获取一个提供用于处理项目正文的方法的对象。</span><span class="sxs-lookup"><span data-stu-id="72c29-230">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="72c29-231">类型：</span><span class="sxs-lookup"><span data-stu-id="72c29-231">Type:</span></span>

*   [<span data-ttu-id="72c29-232">Body</span><span class="sxs-lookup"><span data-stu-id="72c29-232">Body</span></span>](/javascript/api/outlook_1_5/office.body)

##### <a name="requirements"></a><span data-ttu-id="72c29-233">要求</span><span class="sxs-lookup"><span data-stu-id="72c29-233">Requirements</span></span>

|<span data-ttu-id="72c29-234">要求</span><span class="sxs-lookup"><span data-stu-id="72c29-234">Requirement</span></span>| <span data-ttu-id="72c29-235">值</span><span class="sxs-lookup"><span data-stu-id="72c29-235">Value</span></span>|
|---|---|
|[<span data-ttu-id="72c29-236">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="72c29-236">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="72c29-237">1.1</span><span class="sxs-lookup"><span data-stu-id="72c29-237">1.1</span></span>|
|[<span data-ttu-id="72c29-238">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="72c29-238">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="72c29-239">ReadItem</span><span class="sxs-lookup"><span data-stu-id="72c29-239">ReadItem</span></span>|
|[<span data-ttu-id="72c29-240">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="72c29-240">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="72c29-241">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="72c29-241">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="72c29-242">抄送： 数组。 <[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[收件人](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="72c29-242">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="72c29-243">提供对邮件的抄送 (cc) 收件人访问。</span><span class="sxs-lookup"><span data-stu-id="72c29-243">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="72c29-244">对象的类型和访问级别取决于当前项的模式。</span><span class="sxs-lookup"><span data-stu-id="72c29-244">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="72c29-245">阅读模式</span><span class="sxs-lookup"><span data-stu-id="72c29-245">Read mode</span></span>

<span data-ttu-id="72c29-p106">`cc` 属性返回包含邮件的**抄送**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="72c29-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="72c29-248">撰写模式</span><span class="sxs-lookup"><span data-stu-id="72c29-248">Compose mode</span></span>

<span data-ttu-id="72c29-249">`cc`属性返回`Recipients`提供方法用于获取或更新的邮件**Cc**行上的收件人对象。</span><span class="sxs-lookup"><span data-stu-id="72c29-249">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="72c29-250">类型：</span><span class="sxs-lookup"><span data-stu-id="72c29-250">Type:</span></span>

*   <span data-ttu-id="72c29-251">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="72c29-251">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="72c29-252">要求</span><span class="sxs-lookup"><span data-stu-id="72c29-252">Requirements</span></span>

|<span data-ttu-id="72c29-253">要求</span><span class="sxs-lookup"><span data-stu-id="72c29-253">Requirement</span></span>| <span data-ttu-id="72c29-254">值</span><span class="sxs-lookup"><span data-stu-id="72c29-254">Value</span></span>|
|---|---|
|[<span data-ttu-id="72c29-255">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="72c29-255">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="72c29-256">1.0</span><span class="sxs-lookup"><span data-stu-id="72c29-256">1.0</span></span>|
|[<span data-ttu-id="72c29-257">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="72c29-257">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="72c29-258">ReadItem</span><span class="sxs-lookup"><span data-stu-id="72c29-258">ReadItem</span></span>|
|[<span data-ttu-id="72c29-259">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="72c29-259">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="72c29-260">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="72c29-260">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="72c29-261">示例</span><span class="sxs-lookup"><span data-stu-id="72c29-261">Example</span></span>

```
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="72c29-262">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="72c29-262">(nullable) conversationId :String</span></span>

<span data-ttu-id="72c29-263">获取包含特定消息的电子邮件会话的标识符。</span><span class="sxs-lookup"><span data-stu-id="72c29-263">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="72c29-p107">如果在阅读窗体或撰写窗体的回复中激活邮件应用程序，则此属性可以获得一个整数值。如果用户随后更改了回复邮件的主题（若发送回复），则该邮件的对话 ID 将改变且之前获取的值将不适用。</span><span class="sxs-lookup"><span data-stu-id="72c29-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="72c29-p108">对于撰写窗体的新项目，此属性获得一个 null 值。如果用户设置一个主题并保存该项目，`conversationId` 属性将返回一个值。</span><span class="sxs-lookup"><span data-stu-id="72c29-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="72c29-268">类型:</span><span class="sxs-lookup"><span data-stu-id="72c29-268">Type:</span></span>

*   <span data-ttu-id="72c29-269">String</span><span class="sxs-lookup"><span data-stu-id="72c29-269">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="72c29-270">要求</span><span class="sxs-lookup"><span data-stu-id="72c29-270">Requirements</span></span>

|<span data-ttu-id="72c29-271">要求</span><span class="sxs-lookup"><span data-stu-id="72c29-271">Requirement</span></span>| <span data-ttu-id="72c29-272">值</span><span class="sxs-lookup"><span data-stu-id="72c29-272">Value</span></span>|
|---|---|
|[<span data-ttu-id="72c29-273">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="72c29-273">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="72c29-274">1.0</span><span class="sxs-lookup"><span data-stu-id="72c29-274">1.0</span></span>|
|[<span data-ttu-id="72c29-275">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="72c29-275">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="72c29-276">ReadItem</span><span class="sxs-lookup"><span data-stu-id="72c29-276">ReadItem</span></span>|
|[<span data-ttu-id="72c29-277">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="72c29-277">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="72c29-278">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="72c29-278">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="72c29-279">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="72c29-279">dateTimeCreated :Date</span></span>

<span data-ttu-id="72c29-p109">获取项目创建的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="72c29-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="72c29-282">类型：</span><span class="sxs-lookup"><span data-stu-id="72c29-282">Type:</span></span>

*   <span data-ttu-id="72c29-283">日期</span><span class="sxs-lookup"><span data-stu-id="72c29-283">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="72c29-284">要求</span><span class="sxs-lookup"><span data-stu-id="72c29-284">Requirements</span></span>

|<span data-ttu-id="72c29-285">要求</span><span class="sxs-lookup"><span data-stu-id="72c29-285">Requirement</span></span>| <span data-ttu-id="72c29-286">值</span><span class="sxs-lookup"><span data-stu-id="72c29-286">Value</span></span>|
|---|---|
|[<span data-ttu-id="72c29-287">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="72c29-287">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="72c29-288">1.0</span><span class="sxs-lookup"><span data-stu-id="72c29-288">1.0</span></span>|
|[<span data-ttu-id="72c29-289">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="72c29-289">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="72c29-290">ReadItem</span><span class="sxs-lookup"><span data-stu-id="72c29-290">ReadItem</span></span>|
|[<span data-ttu-id="72c29-291">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="72c29-291">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="72c29-292">阅读</span><span class="sxs-lookup"><span data-stu-id="72c29-292">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="72c29-293">示例</span><span class="sxs-lookup"><span data-stu-id="72c29-293">Example</span></span>

```
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="72c29-294">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="72c29-294">dateTimeModified :Date</span></span>

<span data-ttu-id="72c29-p110">获取项目最近一次修改的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="72c29-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="72c29-297">此成员不支持适用于 iOS 的 Outlook 或 Outlook for Android。</span><span class="sxs-lookup"><span data-stu-id="72c29-297">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="72c29-298">类型:</span><span class="sxs-lookup"><span data-stu-id="72c29-298">Type:</span></span>

*   <span data-ttu-id="72c29-299">日期</span><span class="sxs-lookup"><span data-stu-id="72c29-299">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="72c29-300">要求</span><span class="sxs-lookup"><span data-stu-id="72c29-300">Requirements</span></span>

|<span data-ttu-id="72c29-301">要求</span><span class="sxs-lookup"><span data-stu-id="72c29-301">Requirement</span></span>| <span data-ttu-id="72c29-302">值</span><span class="sxs-lookup"><span data-stu-id="72c29-302">Value</span></span>|
|---|---|
|[<span data-ttu-id="72c29-303">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="72c29-303">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="72c29-304">1.0</span><span class="sxs-lookup"><span data-stu-id="72c29-304">1.0</span></span>|
|[<span data-ttu-id="72c29-305">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="72c29-305">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="72c29-306">ReadItem</span><span class="sxs-lookup"><span data-stu-id="72c29-306">ReadItem</span></span>|
|[<span data-ttu-id="72c29-307">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="72c29-307">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="72c29-308">阅读</span><span class="sxs-lookup"><span data-stu-id="72c29-308">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="72c29-309">示例</span><span class="sxs-lookup"><span data-stu-id="72c29-309">Example</span></span>

```
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook15officetime"></a><span data-ttu-id="72c29-310">end :Date|[Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="72c29-310">end :Date|[Time](/javascript/api/outlook_1_5/office.time)</span></span>

<span data-ttu-id="72c29-311">获取或设置约会结束的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="72c29-311">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="72c29-p111">将 `end` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime) 方法将 end 属性值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="72c29-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="72c29-314">阅读模式</span><span class="sxs-lookup"><span data-stu-id="72c29-314">Read mode</span></span>

<span data-ttu-id="72c29-315">`end` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="72c29-315">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="72c29-316">撰写模式</span><span class="sxs-lookup"><span data-stu-id="72c29-316">Compose mode</span></span>

<span data-ttu-id="72c29-317">`end` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="72c29-317">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="72c29-318">使用 [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) 方法设置结束时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="72c29-318">When you use the [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="72c29-319">类型：</span><span class="sxs-lookup"><span data-stu-id="72c29-319">Type:</span></span>

*   <span data-ttu-id="72c29-320">Date | [Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="72c29-320">Date | [Time](/javascript/api/outlook_1_5/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="72c29-321">要求</span><span class="sxs-lookup"><span data-stu-id="72c29-321">Requirements</span></span>

|<span data-ttu-id="72c29-322">要求</span><span class="sxs-lookup"><span data-stu-id="72c29-322">Requirement</span></span>| <span data-ttu-id="72c29-323">值</span><span class="sxs-lookup"><span data-stu-id="72c29-323">Value</span></span>|
|---|---|
|[<span data-ttu-id="72c29-324">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="72c29-324">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="72c29-325">1.0</span><span class="sxs-lookup"><span data-stu-id="72c29-325">1.0</span></span>|
|[<span data-ttu-id="72c29-326">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="72c29-326">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="72c29-327">ReadItem</span><span class="sxs-lookup"><span data-stu-id="72c29-327">ReadItem</span></span>|
|[<span data-ttu-id="72c29-328">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="72c29-328">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="72c29-329">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="72c29-329">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="72c29-330">示例</span><span class="sxs-lookup"><span data-stu-id="72c29-330">Example</span></span>

<span data-ttu-id="72c29-331">以下示例通过使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) 方法，设置撰写模式下约会的结束时间。</span><span class="sxs-lookup"><span data-stu-id="72c29-331">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails"></a><span data-ttu-id="72c29-332">from :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="72c29-332">from :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span></span>

<span data-ttu-id="72c29-p112">获取邮件发件人的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="72c29-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="72c29-p113">`from` 和 [`sender`](#sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="72c29-p113">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="72c29-337">`recipientType`属性`EmailAddressDetails`对象在`from`属性是`undefined`。</span><span class="sxs-lookup"><span data-stu-id="72c29-337">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="72c29-338">类型:</span><span class="sxs-lookup"><span data-stu-id="72c29-338">Type:</span></span>

*   [<span data-ttu-id="72c29-339">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="72c29-339">EmailAddressDetails</span></span>](/javascript/api/outlook_1_5/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="72c29-340">要求</span><span class="sxs-lookup"><span data-stu-id="72c29-340">Requirements</span></span>

|<span data-ttu-id="72c29-341">要求</span><span class="sxs-lookup"><span data-stu-id="72c29-341">Requirement</span></span>| <span data-ttu-id="72c29-342">值</span><span class="sxs-lookup"><span data-stu-id="72c29-342">Value</span></span>|
|---|---|
|[<span data-ttu-id="72c29-343">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="72c29-343">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="72c29-344">1.0</span><span class="sxs-lookup"><span data-stu-id="72c29-344">1.0</span></span>|
|[<span data-ttu-id="72c29-345">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="72c29-345">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="72c29-346">ReadItem</span><span class="sxs-lookup"><span data-stu-id="72c29-346">ReadItem</span></span>|
|[<span data-ttu-id="72c29-347">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="72c29-347">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="72c29-348">阅读</span><span class="sxs-lookup"><span data-stu-id="72c29-348">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="72c29-349">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="72c29-349">internetMessageId :String</span></span>

<span data-ttu-id="72c29-p114">获取电子邮件的 Internet 消息标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="72c29-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="72c29-352">类型：</span><span class="sxs-lookup"><span data-stu-id="72c29-352">Type:</span></span>

*   <span data-ttu-id="72c29-353">String</span><span class="sxs-lookup"><span data-stu-id="72c29-353">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="72c29-354">要求</span><span class="sxs-lookup"><span data-stu-id="72c29-354">Requirements</span></span>

|<span data-ttu-id="72c29-355">要求</span><span class="sxs-lookup"><span data-stu-id="72c29-355">Requirement</span></span>| <span data-ttu-id="72c29-356">值</span><span class="sxs-lookup"><span data-stu-id="72c29-356">Value</span></span>|
|---|---|
|[<span data-ttu-id="72c29-357">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="72c29-357">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="72c29-358">1.0</span><span class="sxs-lookup"><span data-stu-id="72c29-358">1.0</span></span>|
|[<span data-ttu-id="72c29-359">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="72c29-359">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="72c29-360">ReadItem</span><span class="sxs-lookup"><span data-stu-id="72c29-360">ReadItem</span></span>|
|[<span data-ttu-id="72c29-361">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="72c29-361">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="72c29-362">阅读</span><span class="sxs-lookup"><span data-stu-id="72c29-362">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="72c29-363">示例</span><span class="sxs-lookup"><span data-stu-id="72c29-363">Example</span></span>

```
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="72c29-364">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="72c29-364">itemClass :String</span></span>

<span data-ttu-id="72c29-p115">获取选定项目的 Exchange Web 服务项目类。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="72c29-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="72c29-p116">`itemClass` 属性指定所选项目的邮件类别。以下是邮件或约会项目的默认邮件类别。</span><span class="sxs-lookup"><span data-stu-id="72c29-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="72c29-369">类型</span><span class="sxs-lookup"><span data-stu-id="72c29-369">Type</span></span> | <span data-ttu-id="72c29-370">说明</span><span class="sxs-lookup"><span data-stu-id="72c29-370">Description</span></span> | <span data-ttu-id="72c29-371">项目类</span><span class="sxs-lookup"><span data-stu-id="72c29-371">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="72c29-372">约会项目</span><span class="sxs-lookup"><span data-stu-id="72c29-372">Appointment items</span></span> | <span data-ttu-id="72c29-373">这些是项目类 `IPM.Appointment` 或 `IPM.Appointment.Occurence` 的日历项目。</span><span class="sxs-lookup"><span data-stu-id="72c29-373">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| <span data-ttu-id="72c29-374">邮件项目</span><span class="sxs-lookup"><span data-stu-id="72c29-374">Message items</span></span> | <span data-ttu-id="72c29-375">这些项目包括具有默认邮件类别 `IPM.Note` 的电子邮件，以及使用 `IPM.Schedule.Meeting` 作为基础邮件类别的会议请求、响应和取消。</span><span class="sxs-lookup"><span data-stu-id="72c29-375">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="72c29-376">你可以创建用于扩展默认邮件类别的自定义邮件类别，例如，自定义约会邮件类别 `IPM.Appointment.Contoso`。</span><span class="sxs-lookup"><span data-stu-id="72c29-376">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="72c29-377">类型：</span><span class="sxs-lookup"><span data-stu-id="72c29-377">Type:</span></span>

*   <span data-ttu-id="72c29-378">String</span><span class="sxs-lookup"><span data-stu-id="72c29-378">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="72c29-379">要求</span><span class="sxs-lookup"><span data-stu-id="72c29-379">Requirements</span></span>

|<span data-ttu-id="72c29-380">要求</span><span class="sxs-lookup"><span data-stu-id="72c29-380">Requirement</span></span>| <span data-ttu-id="72c29-381">值</span><span class="sxs-lookup"><span data-stu-id="72c29-381">Value</span></span>|
|---|---|
|[<span data-ttu-id="72c29-382">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="72c29-382">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="72c29-383">1.0</span><span class="sxs-lookup"><span data-stu-id="72c29-383">1.0</span></span>|
|[<span data-ttu-id="72c29-384">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="72c29-384">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="72c29-385">ReadItem</span><span class="sxs-lookup"><span data-stu-id="72c29-385">ReadItem</span></span>|
|[<span data-ttu-id="72c29-386">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="72c29-386">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="72c29-387">阅读</span><span class="sxs-lookup"><span data-stu-id="72c29-387">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="72c29-388">示例</span><span class="sxs-lookup"><span data-stu-id="72c29-388">Example</span></span>

```
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="72c29-389">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="72c29-389">(nullable) itemId :String</span></span>

<span data-ttu-id="72c29-p117">获取当前项目的 Exchange Web 服务项目标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="72c29-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="72c29-392">`itemId` 属性返回的标识符与 Exchange Web 服务项目标识符相同。</span><span class="sxs-lookup"><span data-stu-id="72c29-392">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="72c29-393">`itemId`属性不是与 Outlook 条目 ID 或使用 Outlook REST API 的 ID。</span><span class="sxs-lookup"><span data-stu-id="72c29-393">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="72c29-394">使用此值的 REST API 调用之前，应将其转换使用[Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string)。</span><span class="sxs-lookup"><span data-stu-id="72c29-394">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="72c29-395">有关详细信息，请参阅[使用 Outlook REST Api 从 Outlook 外接程序](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id)。</span><span class="sxs-lookup"><span data-stu-id="72c29-395">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="72c29-p119">`itemId` 属性在撰写模式下不可用。如果需要项目标识符，[`saveAsync`](#saveasyncoptions-callback) 方法可用于将项目保存到存储，这将在回调函数的 [`AsyncResult.value`](/javascript/api/office/office.asyncresult) 参数中返回项目标识符。</span><span class="sxs-lookup"><span data-stu-id="72c29-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="72c29-398">类型:</span><span class="sxs-lookup"><span data-stu-id="72c29-398">Type:</span></span>

*   <span data-ttu-id="72c29-399">String</span><span class="sxs-lookup"><span data-stu-id="72c29-399">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="72c29-400">要求</span><span class="sxs-lookup"><span data-stu-id="72c29-400">Requirements</span></span>

|<span data-ttu-id="72c29-401">要求</span><span class="sxs-lookup"><span data-stu-id="72c29-401">Requirement</span></span>| <span data-ttu-id="72c29-402">值</span><span class="sxs-lookup"><span data-stu-id="72c29-402">Value</span></span>|
|---|---|
|[<span data-ttu-id="72c29-403">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="72c29-403">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="72c29-404">1.0</span><span class="sxs-lookup"><span data-stu-id="72c29-404">1.0</span></span>|
|[<span data-ttu-id="72c29-405">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="72c29-405">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="72c29-406">ReadItem</span><span class="sxs-lookup"><span data-stu-id="72c29-406">ReadItem</span></span>|
|[<span data-ttu-id="72c29-407">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="72c29-407">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="72c29-408">阅读</span><span class="sxs-lookup"><span data-stu-id="72c29-408">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="72c29-409">示例</span><span class="sxs-lookup"><span data-stu-id="72c29-409">Example</span></span>

<span data-ttu-id="72c29-p120">以下代码检查项目标识符是否存在。如果 `itemId` 属性返回 `null` 或 `undefined`，则将项目保存到存储，并从异步结果中获取项目标识符。</span><span class="sxs-lookup"><span data-stu-id="72c29-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype"></a><span data-ttu-id="72c29-412">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_5/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="72c29-412">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_5/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="72c29-413">获取实例表示的项的类型。</span><span class="sxs-lookup"><span data-stu-id="72c29-413">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="72c29-414">`itemType` 属性返回其中一个 `ItemType` 枚举值，指示 `item` 对象实例是邮件还是约会。</span><span class="sxs-lookup"><span data-stu-id="72c29-414">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="72c29-415">类型：</span><span class="sxs-lookup"><span data-stu-id="72c29-415">Type:</span></span>

*   [<span data-ttu-id="72c29-416">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="72c29-416">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_5/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="72c29-417">要求</span><span class="sxs-lookup"><span data-stu-id="72c29-417">Requirements</span></span>

|<span data-ttu-id="72c29-418">要求</span><span class="sxs-lookup"><span data-stu-id="72c29-418">Requirement</span></span>| <span data-ttu-id="72c29-419">值</span><span class="sxs-lookup"><span data-stu-id="72c29-419">Value</span></span>|
|---|---|
|[<span data-ttu-id="72c29-420">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="72c29-420">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="72c29-421">1.0</span><span class="sxs-lookup"><span data-stu-id="72c29-421">1.0</span></span>|
|[<span data-ttu-id="72c29-422">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="72c29-422">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="72c29-423">ReadItem</span><span class="sxs-lookup"><span data-stu-id="72c29-423">ReadItem</span></span>|
|[<span data-ttu-id="72c29-424">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="72c29-424">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="72c29-425">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="72c29-425">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="72c29-426">示例</span><span class="sxs-lookup"><span data-stu-id="72c29-426">Example</span></span>

```
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook15officelocation"></a><span data-ttu-id="72c29-427">location :String|[Location](/javascript/api/outlook_1_5/office.location)</span><span class="sxs-lookup"><span data-stu-id="72c29-427">location :String|[Location](/javascript/api/outlook_1_5/office.location)</span></span>

<span data-ttu-id="72c29-428">获取或设置约会的位置。</span><span class="sxs-lookup"><span data-stu-id="72c29-428">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="72c29-429">阅读模式</span><span class="sxs-lookup"><span data-stu-id="72c29-429">Read mode</span></span>

<span data-ttu-id="72c29-430">`location` 属性返回一个包含约会位置的字符串。</span><span class="sxs-lookup"><span data-stu-id="72c29-430">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="72c29-431">撰写模式</span><span class="sxs-lookup"><span data-stu-id="72c29-431">Compose mode</span></span>

<span data-ttu-id="72c29-432">`location` 属性返回一个 `Location` 对象，该对象提供用于获取和设置约会位置的方法。</span><span class="sxs-lookup"><span data-stu-id="72c29-432">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="72c29-433">类型：</span><span class="sxs-lookup"><span data-stu-id="72c29-433">Type:</span></span>

*   <span data-ttu-id="72c29-434">String | [Location](/javascript/api/outlook_1_5/office.location)</span><span class="sxs-lookup"><span data-stu-id="72c29-434">String | [Location](/javascript/api/outlook_1_5/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="72c29-435">要求</span><span class="sxs-lookup"><span data-stu-id="72c29-435">Requirements</span></span>

|<span data-ttu-id="72c29-436">要求</span><span class="sxs-lookup"><span data-stu-id="72c29-436">Requirement</span></span>| <span data-ttu-id="72c29-437">值</span><span class="sxs-lookup"><span data-stu-id="72c29-437">Value</span></span>|
|---|---|
|[<span data-ttu-id="72c29-438">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="72c29-438">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="72c29-439">1.0</span><span class="sxs-lookup"><span data-stu-id="72c29-439">1.0</span></span>|
|[<span data-ttu-id="72c29-440">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="72c29-440">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="72c29-441">ReadItem</span><span class="sxs-lookup"><span data-stu-id="72c29-441">ReadItem</span></span>|
|[<span data-ttu-id="72c29-442">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="72c29-442">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="72c29-443">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="72c29-443">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="72c29-444">示例</span><span class="sxs-lookup"><span data-stu-id="72c29-444">Example</span></span>

```
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="72c29-445">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="72c29-445">normalizedSubject :String</span></span>

<span data-ttu-id="72c29-p121">获取删除了所有前缀（包括 `RE:` 和 `FWD:`）的项目主题。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="72c29-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="72c29-p122">normalizedSubject 属性获取包含由电子邮件程序添加的任何标准前缀（如 `RE:` 和 `FW:`）的项目主题。若要获取包含完整前缀的项目主题，请使用 [`subject`](#subject-stringsubjectjavascriptapioutlook15officesubject) 属性。</span><span class="sxs-lookup"><span data-stu-id="72c29-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook15officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="72c29-450">类型:</span><span class="sxs-lookup"><span data-stu-id="72c29-450">Type:</span></span>

*   <span data-ttu-id="72c29-451">String</span><span class="sxs-lookup"><span data-stu-id="72c29-451">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="72c29-452">要求</span><span class="sxs-lookup"><span data-stu-id="72c29-452">Requirements</span></span>

|<span data-ttu-id="72c29-453">要求</span><span class="sxs-lookup"><span data-stu-id="72c29-453">Requirement</span></span>| <span data-ttu-id="72c29-454">值</span><span class="sxs-lookup"><span data-stu-id="72c29-454">Value</span></span>|
|---|---|
|[<span data-ttu-id="72c29-455">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="72c29-455">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="72c29-456">1.0</span><span class="sxs-lookup"><span data-stu-id="72c29-456">1.0</span></span>|
|[<span data-ttu-id="72c29-457">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="72c29-457">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="72c29-458">ReadItem</span><span class="sxs-lookup"><span data-stu-id="72c29-458">ReadItem</span></span>|
|[<span data-ttu-id="72c29-459">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="72c29-459">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="72c29-460">阅读</span><span class="sxs-lookup"><span data-stu-id="72c29-460">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="72c29-461">示例</span><span class="sxs-lookup"><span data-stu-id="72c29-461">Example</span></span>

```
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook15officenotificationmessages"></a><span data-ttu-id="72c29-462">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_5/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="72c29-462">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_5/office.notificationmessages)</span></span>

<span data-ttu-id="72c29-463">获取项目的通知邮件。</span><span class="sxs-lookup"><span data-stu-id="72c29-463">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="72c29-464">类型：</span><span class="sxs-lookup"><span data-stu-id="72c29-464">Type:</span></span>

*   [<span data-ttu-id="72c29-465">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="72c29-465">NotificationMessages</span></span>](/javascript/api/outlook_1_5/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="72c29-466">要求</span><span class="sxs-lookup"><span data-stu-id="72c29-466">Requirements</span></span>

|<span data-ttu-id="72c29-467">要求</span><span class="sxs-lookup"><span data-stu-id="72c29-467">Requirement</span></span>| <span data-ttu-id="72c29-468">值</span><span class="sxs-lookup"><span data-stu-id="72c29-468">Value</span></span>|
|---|---|
|[<span data-ttu-id="72c29-469">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="72c29-469">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="72c29-470">1.3</span><span class="sxs-lookup"><span data-stu-id="72c29-470">1.3</span></span>|
|[<span data-ttu-id="72c29-471">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="72c29-471">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="72c29-472">ReadItem</span><span class="sxs-lookup"><span data-stu-id="72c29-472">ReadItem</span></span>|
|[<span data-ttu-id="72c29-473">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="72c29-473">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="72c29-474">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="72c29-474">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="72c29-475">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="72c29-475">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="72c29-476">提供对事件的可选与会者的访问。</span><span class="sxs-lookup"><span data-stu-id="72c29-476">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="72c29-477">对象的类型和访问级别取决于当前项的模式。</span><span class="sxs-lookup"><span data-stu-id="72c29-477">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="72c29-478">阅读模式</span><span class="sxs-lookup"><span data-stu-id="72c29-478">Read mode</span></span>

<span data-ttu-id="72c29-479">`optionalAttendees` 属性返回一个数组，其中包含每个可选与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="72c29-479">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="72c29-480">撰写模式</span><span class="sxs-lookup"><span data-stu-id="72c29-480">Compose mode</span></span>

<span data-ttu-id="72c29-481">`optionalAttendees`属性返回`Recipients`对象，它提供用于获取或更新会议的可选与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="72c29-481">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="72c29-482">类型：</span><span class="sxs-lookup"><span data-stu-id="72c29-482">Type:</span></span>

*   <span data-ttu-id="72c29-483">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="72c29-483">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="72c29-484">要求</span><span class="sxs-lookup"><span data-stu-id="72c29-484">Requirements</span></span>

|<span data-ttu-id="72c29-485">要求</span><span class="sxs-lookup"><span data-stu-id="72c29-485">Requirement</span></span>| <span data-ttu-id="72c29-486">值</span><span class="sxs-lookup"><span data-stu-id="72c29-486">Value</span></span>|
|---|---|
|[<span data-ttu-id="72c29-487">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="72c29-487">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="72c29-488">1.0</span><span class="sxs-lookup"><span data-stu-id="72c29-488">1.0</span></span>|
|[<span data-ttu-id="72c29-489">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="72c29-489">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="72c29-490">ReadItem</span><span class="sxs-lookup"><span data-stu-id="72c29-490">ReadItem</span></span>|
|[<span data-ttu-id="72c29-491">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="72c29-491">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="72c29-492">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="72c29-492">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="72c29-493">示例</span><span class="sxs-lookup"><span data-stu-id="72c29-493">Example</span></span>

```
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails"></a><span data-ttu-id="72c29-494">organizer :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="72c29-494">organizer :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span></span>

<span data-ttu-id="72c29-p124">获取指定会议的会议组织者的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="72c29-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="72c29-497">类型:</span><span class="sxs-lookup"><span data-stu-id="72c29-497">Type:</span></span>

*   [<span data-ttu-id="72c29-498">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="72c29-498">EmailAddressDetails</span></span>](/javascript/api/outlook_1_5/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="72c29-499">要求</span><span class="sxs-lookup"><span data-stu-id="72c29-499">Requirements</span></span>

|<span data-ttu-id="72c29-500">要求</span><span class="sxs-lookup"><span data-stu-id="72c29-500">Requirement</span></span>| <span data-ttu-id="72c29-501">值</span><span class="sxs-lookup"><span data-stu-id="72c29-501">Value</span></span>|
|---|---|
|[<span data-ttu-id="72c29-502">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="72c29-502">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="72c29-503">1.0</span><span class="sxs-lookup"><span data-stu-id="72c29-503">1.0</span></span>|
|[<span data-ttu-id="72c29-504">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="72c29-504">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="72c29-505">ReadItem</span><span class="sxs-lookup"><span data-stu-id="72c29-505">ReadItem</span></span>|
|[<span data-ttu-id="72c29-506">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="72c29-506">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="72c29-507">阅读</span><span class="sxs-lookup"><span data-stu-id="72c29-507">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="72c29-508">示例</span><span class="sxs-lookup"><span data-stu-id="72c29-508">Example</span></span>

```
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="72c29-509">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="72c29-509">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="72c29-510">提供对事件的必需与会者的访问。</span><span class="sxs-lookup"><span data-stu-id="72c29-510">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="72c29-511">对象的类型和访问级别取决于当前项的模式。</span><span class="sxs-lookup"><span data-stu-id="72c29-511">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="72c29-512">阅读模式</span><span class="sxs-lookup"><span data-stu-id="72c29-512">Read mode</span></span>

<span data-ttu-id="72c29-513">`requiredAttendees` 属性返回一个数组，其中包含每个必需与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="72c29-513">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="72c29-514">撰写模式</span><span class="sxs-lookup"><span data-stu-id="72c29-514">Compose mode</span></span>

<span data-ttu-id="72c29-515">`requiredAttendees`属性返回`Recipients`对象，它提供用于获取或更新会议的必需的与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="72c29-515">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="72c29-516">类型：</span><span class="sxs-lookup"><span data-stu-id="72c29-516">Type:</span></span>

*   <span data-ttu-id="72c29-517">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="72c29-517">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="72c29-518">要求</span><span class="sxs-lookup"><span data-stu-id="72c29-518">Requirements</span></span>

|<span data-ttu-id="72c29-519">要求</span><span class="sxs-lookup"><span data-stu-id="72c29-519">Requirement</span></span>| <span data-ttu-id="72c29-520">值</span><span class="sxs-lookup"><span data-stu-id="72c29-520">Value</span></span>|
|---|---|
|[<span data-ttu-id="72c29-521">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="72c29-521">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="72c29-522">1.0</span><span class="sxs-lookup"><span data-stu-id="72c29-522">1.0</span></span>|
|[<span data-ttu-id="72c29-523">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="72c29-523">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="72c29-524">ReadItem</span><span class="sxs-lookup"><span data-stu-id="72c29-524">ReadItem</span></span>|
|[<span data-ttu-id="72c29-525">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="72c29-525">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="72c29-526">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="72c29-526">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="72c29-527">示例</span><span class="sxs-lookup"><span data-stu-id="72c29-527">Example</span></span>

```
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails"></a><span data-ttu-id="72c29-528">sender :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="72c29-528">sender :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span></span>

<span data-ttu-id="72c29-p126">获取电子邮件发件人的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="72c29-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="72c29-p127">[`from`](#from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) 和 `sender` 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="72c29-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="72c29-533">`recipientType`属性`EmailAddressDetails`对象在`sender`属性是`undefined`。</span><span class="sxs-lookup"><span data-stu-id="72c29-533">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="72c29-534">类型:</span><span class="sxs-lookup"><span data-stu-id="72c29-534">Type:</span></span>

*   [<span data-ttu-id="72c29-535">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="72c29-535">EmailAddressDetails</span></span>](/javascript/api/outlook_1_5/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="72c29-536">要求</span><span class="sxs-lookup"><span data-stu-id="72c29-536">Requirements</span></span>

|<span data-ttu-id="72c29-537">要求</span><span class="sxs-lookup"><span data-stu-id="72c29-537">Requirement</span></span>| <span data-ttu-id="72c29-538">值</span><span class="sxs-lookup"><span data-stu-id="72c29-538">Value</span></span>|
|---|---|
|[<span data-ttu-id="72c29-539">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="72c29-539">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="72c29-540">1.0</span><span class="sxs-lookup"><span data-stu-id="72c29-540">1.0</span></span>|
|[<span data-ttu-id="72c29-541">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="72c29-541">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="72c29-542">ReadItem</span><span class="sxs-lookup"><span data-stu-id="72c29-542">ReadItem</span></span>|
|[<span data-ttu-id="72c29-543">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="72c29-543">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="72c29-544">阅读</span><span class="sxs-lookup"><span data-stu-id="72c29-544">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="72c29-545">示例</span><span class="sxs-lookup"><span data-stu-id="72c29-545">Example</span></span>

```
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook15officetime"></a><span data-ttu-id="72c29-546">start :Date|[Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="72c29-546">start :Date|[Time](/javascript/api/outlook_1_5/office.time)</span></span>

<span data-ttu-id="72c29-547">获取或设置约会开始的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="72c29-547">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="72c29-p128">将 `start` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime) 方法将该值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="72c29-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="72c29-550">阅读模式</span><span class="sxs-lookup"><span data-stu-id="72c29-550">Read mode</span></span>

<span data-ttu-id="72c29-551">`start` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="72c29-551">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="72c29-552">撰写模式</span><span class="sxs-lookup"><span data-stu-id="72c29-552">Compose mode</span></span>

<span data-ttu-id="72c29-553">`start` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="72c29-553">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="72c29-554">使用 [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) 方法设置开始时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="72c29-554">When you use the [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="72c29-555">类型：</span><span class="sxs-lookup"><span data-stu-id="72c29-555">Type:</span></span>

*   <span data-ttu-id="72c29-556">Date | [Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="72c29-556">Date | [Time](/javascript/api/outlook_1_5/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="72c29-557">要求</span><span class="sxs-lookup"><span data-stu-id="72c29-557">Requirements</span></span>

|<span data-ttu-id="72c29-558">要求</span><span class="sxs-lookup"><span data-stu-id="72c29-558">Requirement</span></span>| <span data-ttu-id="72c29-559">值</span><span class="sxs-lookup"><span data-stu-id="72c29-559">Value</span></span>|
|---|---|
|[<span data-ttu-id="72c29-560">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="72c29-560">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="72c29-561">1.0</span><span class="sxs-lookup"><span data-stu-id="72c29-561">1.0</span></span>|
|[<span data-ttu-id="72c29-562">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="72c29-562">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="72c29-563">ReadItem</span><span class="sxs-lookup"><span data-stu-id="72c29-563">ReadItem</span></span>|
|[<span data-ttu-id="72c29-564">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="72c29-564">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="72c29-565">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="72c29-565">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="72c29-566">示例</span><span class="sxs-lookup"><span data-stu-id="72c29-566">Example</span></span>

<span data-ttu-id="72c29-567">以下示例通过使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) 方法，设置撰写模式下约会的开始时间。</span><span class="sxs-lookup"><span data-stu-id="72c29-567">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

####  <a name="subject-stringsubjectjavascriptapioutlook15officesubject"></a><span data-ttu-id="72c29-568">subject :String|[Subject](/javascript/api/outlook_1_5/office.subject)</span><span class="sxs-lookup"><span data-stu-id="72c29-568">subject :String|[Subject](/javascript/api/outlook_1_5/office.subject)</span></span>

<span data-ttu-id="72c29-569">获取或设置显示在项目的主题字段中的说明。</span><span class="sxs-lookup"><span data-stu-id="72c29-569">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="72c29-570">`subject` 属性获取或设置由电子邮件服务器发送项目时的整个主题。</span><span class="sxs-lookup"><span data-stu-id="72c29-570">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="72c29-571">阅读模式</span><span class="sxs-lookup"><span data-stu-id="72c29-571">Read mode</span></span>

<span data-ttu-id="72c29-p129">`subject` 属性返回一个字符串。使用 [`normalizedSubject`](#normalizedsubject-string) 属性获取不带任何前导前缀（如 `RE:` 和 `FW:`）的主题。</span><span class="sxs-lookup"><span data-stu-id="72c29-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="72c29-574">撰写模式</span><span class="sxs-lookup"><span data-stu-id="72c29-574">Compose mode</span></span>

<span data-ttu-id="72c29-575">`subject` 属性返回一个 `Subject` 对象，该对象提供用于获取和设置主题的方法。</span><span class="sxs-lookup"><span data-stu-id="72c29-575">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="72c29-576">类型：</span><span class="sxs-lookup"><span data-stu-id="72c29-576">Type:</span></span>

*   <span data-ttu-id="72c29-577">String | [Subject](/javascript/api/outlook_1_5/office.subject)</span><span class="sxs-lookup"><span data-stu-id="72c29-577">String | [Subject](/javascript/api/outlook_1_5/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="72c29-578">要求</span><span class="sxs-lookup"><span data-stu-id="72c29-578">Requirements</span></span>

|<span data-ttu-id="72c29-579">要求</span><span class="sxs-lookup"><span data-stu-id="72c29-579">Requirement</span></span>| <span data-ttu-id="72c29-580">值</span><span class="sxs-lookup"><span data-stu-id="72c29-580">Value</span></span>|
|---|---|
|[<span data-ttu-id="72c29-581">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="72c29-581">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="72c29-582">1.0</span><span class="sxs-lookup"><span data-stu-id="72c29-582">1.0</span></span>|
|[<span data-ttu-id="72c29-583">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="72c29-583">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="72c29-584">ReadItem</span><span class="sxs-lookup"><span data-stu-id="72c29-584">ReadItem</span></span>|
|[<span data-ttu-id="72c29-585">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="72c29-585">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="72c29-586">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="72c29-586">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="72c29-587">到： 数组。 <[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[收件人](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="72c29-587">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="72c29-588">提供对邮件的**到**行上的收件人访问。</span><span class="sxs-lookup"><span data-stu-id="72c29-588">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="72c29-589">对象的类型和访问级别取决于当前项的模式。</span><span class="sxs-lookup"><span data-stu-id="72c29-589">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="72c29-590">阅读模式</span><span class="sxs-lookup"><span data-stu-id="72c29-590">Read mode</span></span>

<span data-ttu-id="72c29-p131">`to` 属性返回包含邮件的**收件人**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="72c29-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="72c29-593">撰写模式</span><span class="sxs-lookup"><span data-stu-id="72c29-593">Compose mode</span></span>

<span data-ttu-id="72c29-594">`to`属性返回`Recipients`提供方法用于获取或更新邮件的**到**行上的收件人对象。</span><span class="sxs-lookup"><span data-stu-id="72c29-594">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="72c29-595">类型：</span><span class="sxs-lookup"><span data-stu-id="72c29-595">Type:</span></span>

*   <span data-ttu-id="72c29-596">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="72c29-596">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="72c29-597">要求</span><span class="sxs-lookup"><span data-stu-id="72c29-597">Requirements</span></span>

|<span data-ttu-id="72c29-598">要求</span><span class="sxs-lookup"><span data-stu-id="72c29-598">Requirement</span></span>| <span data-ttu-id="72c29-599">值</span><span class="sxs-lookup"><span data-stu-id="72c29-599">Value</span></span>|
|---|---|
|[<span data-ttu-id="72c29-600">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="72c29-600">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="72c29-601">1.0</span><span class="sxs-lookup"><span data-stu-id="72c29-601">1.0</span></span>|
|[<span data-ttu-id="72c29-602">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="72c29-602">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="72c29-603">ReadItem</span><span class="sxs-lookup"><span data-stu-id="72c29-603">ReadItem</span></span>|
|[<span data-ttu-id="72c29-604">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="72c29-604">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="72c29-605">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="72c29-605">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="72c29-606">示例</span><span class="sxs-lookup"><span data-stu-id="72c29-606">Example</span></span>

```
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="72c29-607">方法</span><span class="sxs-lookup"><span data-stu-id="72c29-607">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="72c29-608">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="72c29-608">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="72c29-609">将文件作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="72c29-609">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="72c29-610">`addFileAttachmentAsync` 方法在指定的 URI 上载文件并将其附加到撰写窗体中的项目。</span><span class="sxs-lookup"><span data-stu-id="72c29-610">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="72c29-611">你随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="72c29-611">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="72c29-612">参数：</span><span class="sxs-lookup"><span data-stu-id="72c29-612">Parameters:</span></span>

|<span data-ttu-id="72c29-613">名称</span><span class="sxs-lookup"><span data-stu-id="72c29-613">Name</span></span>| <span data-ttu-id="72c29-614">类型</span><span class="sxs-lookup"><span data-stu-id="72c29-614">Type</span></span>| <span data-ttu-id="72c29-615">属性</span><span class="sxs-lookup"><span data-stu-id="72c29-615">Attributes</span></span>| <span data-ttu-id="72c29-616">说明</span><span class="sxs-lookup"><span data-stu-id="72c29-616">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="72c29-617">字符串</span><span class="sxs-lookup"><span data-stu-id="72c29-617">String</span></span>||<span data-ttu-id="72c29-p132">提供附加到邮件或约会的文件的位置的 URI。最大长度为 2048 个字符。</span><span class="sxs-lookup"><span data-stu-id="72c29-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="72c29-620">字符串</span><span class="sxs-lookup"><span data-stu-id="72c29-620">String</span></span>||<span data-ttu-id="72c29-p133">在附件上载过程中显示的附件名称。最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="72c29-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="72c29-623">对象</span><span class="sxs-lookup"><span data-stu-id="72c29-623">Object</span></span>| <span data-ttu-id="72c29-624">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="72c29-624">&lt;optional&gt;</span></span>|<span data-ttu-id="72c29-625">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="72c29-625">An object literal that contains one or more of the following properties.</span></span>|
| `options.asyncContext` | <span data-ttu-id="72c29-626">对象</span><span class="sxs-lookup"><span data-stu-id="72c29-626">Object</span></span> | <span data-ttu-id="72c29-627">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="72c29-627">&lt;optional&gt;</span></span> | <span data-ttu-id="72c29-628">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="72c29-628">Developers can provide any object they wish to access in the callback method.</span></span> |
| `options.isInline` | <span data-ttu-id="72c29-629">布尔值</span><span class="sxs-lookup"><span data-stu-id="72c29-629">Boolean</span></span> | <span data-ttu-id="72c29-630">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="72c29-630">&lt;optional&gt;</span></span> | <span data-ttu-id="72c29-631">如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="72c29-631">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
|`callback`| <span data-ttu-id="72c29-632">函数</span><span class="sxs-lookup"><span data-stu-id="72c29-632">function</span></span>| <span data-ttu-id="72c29-633">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="72c29-633">&lt;optional&gt;</span></span>|<span data-ttu-id="72c29-634">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="72c29-634">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="72c29-635">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="72c29-635">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="72c29-636">如果上传附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="72c29-636">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="72c29-637">错误</span><span class="sxs-lookup"><span data-stu-id="72c29-637">Errors</span></span>

| <span data-ttu-id="72c29-638">错误代码</span><span class="sxs-lookup"><span data-stu-id="72c29-638">Error code</span></span> | <span data-ttu-id="72c29-639">说明</span><span class="sxs-lookup"><span data-stu-id="72c29-639">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="72c29-640">附件大小超过了允许的大小。</span><span class="sxs-lookup"><span data-stu-id="72c29-640">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="72c29-641">该附件的扩展名不是允许的扩展名。</span><span class="sxs-lookup"><span data-stu-id="72c29-641">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="72c29-642">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="72c29-642">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="72c29-643">要求</span><span class="sxs-lookup"><span data-stu-id="72c29-643">Requirements</span></span>

|<span data-ttu-id="72c29-644">要求</span><span class="sxs-lookup"><span data-stu-id="72c29-644">Requirement</span></span>| <span data-ttu-id="72c29-645">值</span><span class="sxs-lookup"><span data-stu-id="72c29-645">Value</span></span>|
|---|---|
|[<span data-ttu-id="72c29-646">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="72c29-646">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="72c29-647">1.1</span><span class="sxs-lookup"><span data-stu-id="72c29-647">1.1</span></span>|
|[<span data-ttu-id="72c29-648">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="72c29-648">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="72c29-649">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="72c29-649">ReadWriteItem</span></span>|
|[<span data-ttu-id="72c29-650">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="72c29-650">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="72c29-651">撰写</span><span class="sxs-lookup"><span data-stu-id="72c29-651">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="72c29-652">示例</span><span class="sxs-lookup"><span data-stu-id="72c29-652">Examples</span></span>

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

<span data-ttu-id="72c29-653">以下示例将图像文件添加为内联附件，并在邮件正文中引用该附件。</span><span class="sxs-lookup"><span data-stu-id="72c29-653">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="72c29-654">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="72c29-654">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="72c29-655">将 Exchange 项目（如邮件）作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="72c29-655">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="72c29-p134">`addItemAttachmentAsync` 方法将包含指定 Exchange 标识符的项目附加到撰写窗体中的项目。如果指定一个回调方法，此方法使用 `asyncResult` 参数调用，该参数包含一个附件标识符或代码，指示附加项目过程中出现的任何错误。可以使用 `options` 参数将状态信息传递给回调方法（如果需要）。</span><span class="sxs-lookup"><span data-stu-id="72c29-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="72c29-659">你随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="72c29-659">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="72c29-660">如果您 Office 加载项运行的 Outlook Web App 中`addItemAttachmentAsync`方法可以附加到正在编辑; 的项目之外的项目的项目但是，这不受支持，并且不建议使用。</span><span class="sxs-lookup"><span data-stu-id="72c29-660">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="72c29-661">参数：</span><span class="sxs-lookup"><span data-stu-id="72c29-661">Parameters:</span></span>

|<span data-ttu-id="72c29-662">名称</span><span class="sxs-lookup"><span data-stu-id="72c29-662">Name</span></span>| <span data-ttu-id="72c29-663">类型</span><span class="sxs-lookup"><span data-stu-id="72c29-663">Type</span></span>| <span data-ttu-id="72c29-664">属性</span><span class="sxs-lookup"><span data-stu-id="72c29-664">Attributes</span></span>| <span data-ttu-id="72c29-665">说明</span><span class="sxs-lookup"><span data-stu-id="72c29-665">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="72c29-666">字符串</span><span class="sxs-lookup"><span data-stu-id="72c29-666">String</span></span>||<span data-ttu-id="72c29-p135">要附加的项目的 Exchange 标识符。最大长度为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="72c29-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="72c29-669">字符串</span><span class="sxs-lookup"><span data-stu-id="72c29-669">String</span></span>||<span data-ttu-id="72c29-p136">要附加的项目的主题。最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="72c29-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="72c29-672">对象</span><span class="sxs-lookup"><span data-stu-id="72c29-672">Object</span></span>| <span data-ttu-id="72c29-673">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="72c29-673">&lt;optional&gt;</span></span>|<span data-ttu-id="72c29-674">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="72c29-674">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="72c29-675">对象</span><span class="sxs-lookup"><span data-stu-id="72c29-675">Object</span></span>| <span data-ttu-id="72c29-676">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="72c29-676">&lt;optional&gt;</span></span>|<span data-ttu-id="72c29-677">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="72c29-677">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="72c29-678">函数</span><span class="sxs-lookup"><span data-stu-id="72c29-678">function</span></span>| <span data-ttu-id="72c29-679">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="72c29-679">&lt;optional&gt;</span></span>|<span data-ttu-id="72c29-680">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="72c29-680">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="72c29-681">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="72c29-681">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="72c29-682">如果添加附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="72c29-682">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="72c29-683">错误</span><span class="sxs-lookup"><span data-stu-id="72c29-683">Errors</span></span>

| <span data-ttu-id="72c29-684">错误代码</span><span class="sxs-lookup"><span data-stu-id="72c29-684">Error code</span></span> | <span data-ttu-id="72c29-685">说明</span><span class="sxs-lookup"><span data-stu-id="72c29-685">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="72c29-686">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="72c29-686">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="72c29-687">要求</span><span class="sxs-lookup"><span data-stu-id="72c29-687">Requirements</span></span>

|<span data-ttu-id="72c29-688">要求</span><span class="sxs-lookup"><span data-stu-id="72c29-688">Requirement</span></span>| <span data-ttu-id="72c29-689">值</span><span class="sxs-lookup"><span data-stu-id="72c29-689">Value</span></span>|
|---|---|
|[<span data-ttu-id="72c29-690">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="72c29-690">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="72c29-691">1.1</span><span class="sxs-lookup"><span data-stu-id="72c29-691">1.1</span></span>|
|[<span data-ttu-id="72c29-692">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="72c29-692">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="72c29-693">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="72c29-693">ReadWriteItem</span></span>|
|[<span data-ttu-id="72c29-694">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="72c29-694">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="72c29-695">撰写</span><span class="sxs-lookup"><span data-stu-id="72c29-695">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="72c29-696">示例</span><span class="sxs-lookup"><span data-stu-id="72c29-696">Example</span></span>

<span data-ttu-id="72c29-697">以下示例将现有的 Outlook 项目添加为名为 `My Attachment` 的附件。</span><span class="sxs-lookup"><span data-stu-id="72c29-697">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

####  <a name="close"></a><span data-ttu-id="72c29-698">close()</span><span class="sxs-lookup"><span data-stu-id="72c29-698">close()</span></span>

<span data-ttu-id="72c29-699">关闭当前正在撰写的项目。</span><span class="sxs-lookup"><span data-stu-id="72c29-699">Closes the current item that is being composed.</span></span>

<span data-ttu-id="72c29-p137">`close` 方法的行为取决于要撰写的项目的当前状态。如果项目具有未保存的更改，客户端将提示用户保存、放弃或取消关闭操作。</span><span class="sxs-lookup"><span data-stu-id="72c29-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="72c29-702">在 Outlook 中的 web，如果该项目是约会和它之前已保存使用在`saveAsync`、 提示用户保存、 放弃，或取消，即使未发生更改由于项目上次保存。</span><span class="sxs-lookup"><span data-stu-id="72c29-702">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="72c29-703">在 Outlook 桌面客户端中，如果邮件是内联答复，`close` 方法不起作用。</span><span class="sxs-lookup"><span data-stu-id="72c29-703">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="72c29-704">要求</span><span class="sxs-lookup"><span data-stu-id="72c29-704">Requirements</span></span>

|<span data-ttu-id="72c29-705">要求</span><span class="sxs-lookup"><span data-stu-id="72c29-705">Requirement</span></span>| <span data-ttu-id="72c29-706">值</span><span class="sxs-lookup"><span data-stu-id="72c29-706">Value</span></span>|
|---|---|
|[<span data-ttu-id="72c29-707">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="72c29-707">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="72c29-708">1.3</span><span class="sxs-lookup"><span data-stu-id="72c29-708">1.3</span></span>|
|[<span data-ttu-id="72c29-709">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="72c29-709">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="72c29-710">受限</span><span class="sxs-lookup"><span data-stu-id="72c29-710">Restricted</span></span>|
|[<span data-ttu-id="72c29-711">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="72c29-711">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="72c29-712">撰写</span><span class="sxs-lookup"><span data-stu-id="72c29-712">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="72c29-713">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="72c29-713">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="72c29-714">显示答复窗体，其中包括所选邮件的发件人和所有收件人或所选约会的组织者和所有与会者。</span><span class="sxs-lookup"><span data-stu-id="72c29-714">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="72c29-715">在适用于 iOS 的 Outlook 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="72c29-715">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="72c29-716">在 Outlook Web App 中，答复窗体显示为包含 3 列视图的弹出式窗体以及包含 2 列或 1 列视图的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="72c29-716">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="72c29-717">如果任意字符串参数超出其限制，`displayReplyAllForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="72c29-717">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="72c29-p138">当在 `formData.attachments` 参数中指定附件时，Outlook 和 Outlook Web App 尝试下载所有附件并将其附加到答复窗体。如果无法添加任何附件，则在窗体 UI 中显示错误。如果这不可能，则不引发错误消息。</span><span class="sxs-lookup"><span data-stu-id="72c29-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="72c29-721">参数：</span><span class="sxs-lookup"><span data-stu-id="72c29-721">Parameters:</span></span>

| <span data-ttu-id="72c29-722">名称</span><span class="sxs-lookup"><span data-stu-id="72c29-722">Name</span></span> | <span data-ttu-id="72c29-723">类型</span><span class="sxs-lookup"><span data-stu-id="72c29-723">Type</span></span> | <span data-ttu-id="72c29-724">属性</span><span class="sxs-lookup"><span data-stu-id="72c29-724">Attributes</span></span> | <span data-ttu-id="72c29-725">说明</span><span class="sxs-lookup"><span data-stu-id="72c29-725">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="72c29-726">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="72c29-726">String &#124; Object</span></span>| |<span data-ttu-id="72c29-p139">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="72c29-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="72c29-729">**或**</span><span class="sxs-lookup"><span data-stu-id="72c29-729">**OR**</span></span><br/><span data-ttu-id="72c29-p140">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="72c29-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="72c29-732">字符串</span><span class="sxs-lookup"><span data-stu-id="72c29-732">String</span></span> | <span data-ttu-id="72c29-733">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="72c29-733">&lt;optional&gt;</span></span> | <span data-ttu-id="72c29-p141">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="72c29-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="72c29-736">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="72c29-736">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="72c29-737">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="72c29-737">&lt;optional&gt;</span></span> | <span data-ttu-id="72c29-738">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="72c29-738">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="72c29-739">字符串</span><span class="sxs-lookup"><span data-stu-id="72c29-739">String</span></span> | | <span data-ttu-id="72c29-p142">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="72c29-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="72c29-742">字符串</span><span class="sxs-lookup"><span data-stu-id="72c29-742">String</span></span> | | <span data-ttu-id="72c29-743">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="72c29-743">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="72c29-744">String</span><span class="sxs-lookup"><span data-stu-id="72c29-744">String</span></span> | | <span data-ttu-id="72c29-p143">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="72c29-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="72c29-747">Boolean</span><span class="sxs-lookup"><span data-stu-id="72c29-747">Boolean</span></span> | | <span data-ttu-id="72c29-p144">仅在将 `type` 设置为 `file` 时使用。如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="72c29-p144">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="72c29-750">String</span><span class="sxs-lookup"><span data-stu-id="72c29-750">String</span></span> | | <span data-ttu-id="72c29-p145">仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="72c29-p145">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="72c29-754">函数</span><span class="sxs-lookup"><span data-stu-id="72c29-754">function</span></span> | <span data-ttu-id="72c29-755">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="72c29-755">&lt;optional&gt;</span></span> | <span data-ttu-id="72c29-756">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="72c29-756">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="72c29-757">要求</span><span class="sxs-lookup"><span data-stu-id="72c29-757">Requirements</span></span>

|<span data-ttu-id="72c29-758">要求</span><span class="sxs-lookup"><span data-stu-id="72c29-758">Requirement</span></span>| <span data-ttu-id="72c29-759">值</span><span class="sxs-lookup"><span data-stu-id="72c29-759">Value</span></span>|
|---|---|
|[<span data-ttu-id="72c29-760">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="72c29-760">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="72c29-761">1.0</span><span class="sxs-lookup"><span data-stu-id="72c29-761">1.0</span></span>|
|[<span data-ttu-id="72c29-762">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="72c29-762">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="72c29-763">ReadItem</span><span class="sxs-lookup"><span data-stu-id="72c29-763">ReadItem</span></span>|
|[<span data-ttu-id="72c29-764">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="72c29-764">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="72c29-765">阅读</span><span class="sxs-lookup"><span data-stu-id="72c29-765">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="72c29-766">示例</span><span class="sxs-lookup"><span data-stu-id="72c29-766">Examples</span></span>

<span data-ttu-id="72c29-767">以下代码将一个字符串传递到 `displayReplyAllForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="72c29-767">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="72c29-768">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="72c29-768">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="72c29-769">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="72c29-769">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="72c29-770">使用正文和文件附件答复。</span><span class="sxs-lookup"><span data-stu-id="72c29-770">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="72c29-771">使用正文和项目附件答复。</span><span class="sxs-lookup"><span data-stu-id="72c29-771">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="72c29-772">使用正文、文件附件、项目附件和回调答复。</span><span class="sxs-lookup"><span data-stu-id="72c29-772">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="72c29-773">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="72c29-773">displayReplyForm(formData)</span></span>

<span data-ttu-id="72c29-774">显示答复窗体，其中仅包括所选邮件的发件人或所选约会的组织者。</span><span class="sxs-lookup"><span data-stu-id="72c29-774">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="72c29-775">在适用于 iOS 的 Outlook 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="72c29-775">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="72c29-776">在 Outlook Web App 中，答复窗体显示为包含 3 列视图的弹出式窗体以及包含 2 列或 1 列视图的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="72c29-776">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="72c29-777">如果任意字符串参数超出其限制，`displayReplyForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="72c29-777">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="72c29-p146">当在 `formData.attachments` 参数中指定附件时，Outlook 和 Outlook Web App 尝试下载所有附件并将其附加到答复窗体。如果无法添加任何附件，则在窗体 UI 中显示错误。如果这不可能，则不引发错误消息。</span><span class="sxs-lookup"><span data-stu-id="72c29-p146">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="72c29-781">参数：</span><span class="sxs-lookup"><span data-stu-id="72c29-781">Parameters:</span></span>

| <span data-ttu-id="72c29-782">名称</span><span class="sxs-lookup"><span data-stu-id="72c29-782">Name</span></span> | <span data-ttu-id="72c29-783">类型</span><span class="sxs-lookup"><span data-stu-id="72c29-783">Type</span></span> | <span data-ttu-id="72c29-784">属性</span><span class="sxs-lookup"><span data-stu-id="72c29-784">Attributes</span></span> | <span data-ttu-id="72c29-785">说明</span><span class="sxs-lookup"><span data-stu-id="72c29-785">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="72c29-786">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="72c29-786">String &#124; Object</span></span>| | <span data-ttu-id="72c29-p147">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="72c29-p147">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="72c29-789">**或**</span><span class="sxs-lookup"><span data-stu-id="72c29-789">**OR**</span></span><br/><span data-ttu-id="72c29-p148">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="72c29-p148">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="72c29-792">字符串</span><span class="sxs-lookup"><span data-stu-id="72c29-792">String</span></span> | <span data-ttu-id="72c29-793">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="72c29-793">&lt;optional&gt;</span></span> | <span data-ttu-id="72c29-p149">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="72c29-p149">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="72c29-796">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="72c29-796">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="72c29-797">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="72c29-797">&lt;optional&gt;</span></span> | <span data-ttu-id="72c29-798">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="72c29-798">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="72c29-799">字符串</span><span class="sxs-lookup"><span data-stu-id="72c29-799">String</span></span> | | <span data-ttu-id="72c29-p150">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="72c29-p150">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="72c29-802">字符串</span><span class="sxs-lookup"><span data-stu-id="72c29-802">String</span></span> | | <span data-ttu-id="72c29-803">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="72c29-803">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="72c29-804">String</span><span class="sxs-lookup"><span data-stu-id="72c29-804">String</span></span> | | <span data-ttu-id="72c29-p151">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="72c29-p151">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="72c29-807">Boolean</span><span class="sxs-lookup"><span data-stu-id="72c29-807">Boolean</span></span> | | <span data-ttu-id="72c29-p152">仅在将 `type` 设置为 `file` 时使用。如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="72c29-p152">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="72c29-810">String</span><span class="sxs-lookup"><span data-stu-id="72c29-810">String</span></span> | | <span data-ttu-id="72c29-p153">仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="72c29-p153">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="72c29-814">函数</span><span class="sxs-lookup"><span data-stu-id="72c29-814">function</span></span> | <span data-ttu-id="72c29-815">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="72c29-815">&lt;optional&gt;</span></span> | <span data-ttu-id="72c29-816">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="72c29-816">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="72c29-817">要求</span><span class="sxs-lookup"><span data-stu-id="72c29-817">Requirements</span></span>

|<span data-ttu-id="72c29-818">要求</span><span class="sxs-lookup"><span data-stu-id="72c29-818">Requirement</span></span>| <span data-ttu-id="72c29-819">值</span><span class="sxs-lookup"><span data-stu-id="72c29-819">Value</span></span>|
|---|---|
|[<span data-ttu-id="72c29-820">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="72c29-820">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="72c29-821">1.0</span><span class="sxs-lookup"><span data-stu-id="72c29-821">1.0</span></span>|
|[<span data-ttu-id="72c29-822">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="72c29-822">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="72c29-823">ReadItem</span><span class="sxs-lookup"><span data-stu-id="72c29-823">ReadItem</span></span>|
|[<span data-ttu-id="72c29-824">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="72c29-824">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="72c29-825">阅读</span><span class="sxs-lookup"><span data-stu-id="72c29-825">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="72c29-826">示例</span><span class="sxs-lookup"><span data-stu-id="72c29-826">Examples</span></span>

<span data-ttu-id="72c29-827">以下代码将一个字符串传递到 `displayReplyForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="72c29-827">The following code passes a string to the `displayReplyForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="72c29-828">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="72c29-828">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="72c29-829">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="72c29-829">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="72c29-830">使用正文和文件附件答复。</span><span class="sxs-lookup"><span data-stu-id="72c29-830">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="72c29-831">使用正文和项目附件答复。</span><span class="sxs-lookup"><span data-stu-id="72c29-831">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="72c29-832">使用正文、文件附件、项目附件和回调答复。</span><span class="sxs-lookup"><span data-stu-id="72c29-832">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook15officeentities"></a><span data-ttu-id="72c29-833">getEntities() → {[Entities](/javascript/api/outlook_1_5/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="72c29-833">getEntities() → {[Entities](/javascript/api/outlook_1_5/office.entities)}</span></span>

<span data-ttu-id="72c29-834">获取在选定的项的正文中找到的实体。</span><span class="sxs-lookup"><span data-stu-id="72c29-834">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="72c29-835">在适用于 iOS 的 Outlook 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="72c29-835">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="72c29-836">要求</span><span class="sxs-lookup"><span data-stu-id="72c29-836">Requirements</span></span>

|<span data-ttu-id="72c29-837">要求</span><span class="sxs-lookup"><span data-stu-id="72c29-837">Requirement</span></span>| <span data-ttu-id="72c29-838">值</span><span class="sxs-lookup"><span data-stu-id="72c29-838">Value</span></span>|
|---|---|
|[<span data-ttu-id="72c29-839">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="72c29-839">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="72c29-840">1.0</span><span class="sxs-lookup"><span data-stu-id="72c29-840">1.0</span></span>|
|[<span data-ttu-id="72c29-841">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="72c29-841">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="72c29-842">ReadItem</span><span class="sxs-lookup"><span data-stu-id="72c29-842">ReadItem</span></span>|
|[<span data-ttu-id="72c29-843">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="72c29-843">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="72c29-844">阅读</span><span class="sxs-lookup"><span data-stu-id="72c29-844">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="72c29-845">返回：</span><span class="sxs-lookup"><span data-stu-id="72c29-845">Returns:</span></span>

<span data-ttu-id="72c29-846">类型：[Entities](/javascript/api/outlook_1_5/office.entities)</span><span class="sxs-lookup"><span data-stu-id="72c29-846">Type: [Entities](/javascript/api/outlook_1_5/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="72c29-847">示例</span><span class="sxs-lookup"><span data-stu-id="72c29-847">Example</span></span>

<span data-ttu-id="72c29-848">以下示例访问当前项的主体中的联系人实体。</span><span class="sxs-lookup"><span data-stu-id="72c29-848">The following example accesses the contacts entities in the current item's body.</span></span>

```
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion"></a><span data-ttu-id="72c29-849">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="72c29-849">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span></span>

<span data-ttu-id="72c29-850">获取选定的项的正文中找到的指定的实体类型的所有实体的数组。</span><span class="sxs-lookup"><span data-stu-id="72c29-850">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="72c29-851">在适用于 iOS 的 Outlook 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="72c29-851">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="72c29-852">参数：</span><span class="sxs-lookup"><span data-stu-id="72c29-852">Parameters:</span></span>

|<span data-ttu-id="72c29-853">名称</span><span class="sxs-lookup"><span data-stu-id="72c29-853">Name</span></span>| <span data-ttu-id="72c29-854">类型</span><span class="sxs-lookup"><span data-stu-id="72c29-854">Type</span></span>| <span data-ttu-id="72c29-855">说明</span><span class="sxs-lookup"><span data-stu-id="72c29-855">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="72c29-856">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="72c29-856">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_5/office.mailboxenums.entitytype)|<span data-ttu-id="72c29-857">EntityType 枚举值之一。</span><span class="sxs-lookup"><span data-stu-id="72c29-857">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="72c29-858">要求</span><span class="sxs-lookup"><span data-stu-id="72c29-858">Requirements</span></span>

|<span data-ttu-id="72c29-859">要求</span><span class="sxs-lookup"><span data-stu-id="72c29-859">Requirement</span></span>| <span data-ttu-id="72c29-860">值</span><span class="sxs-lookup"><span data-stu-id="72c29-860">Value</span></span>|
|---|---|
|[<span data-ttu-id="72c29-861">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="72c29-861">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="72c29-862">1.0</span><span class="sxs-lookup"><span data-stu-id="72c29-862">1.0</span></span>|
|[<span data-ttu-id="72c29-863">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="72c29-863">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="72c29-864">受限</span><span class="sxs-lookup"><span data-stu-id="72c29-864">Restricted</span></span>|
|[<span data-ttu-id="72c29-865">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="72c29-865">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="72c29-866">阅读</span><span class="sxs-lookup"><span data-stu-id="72c29-866">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="72c29-867">返回：</span><span class="sxs-lookup"><span data-stu-id="72c29-867">Returns:</span></span>

<span data-ttu-id="72c29-868">如果在 `entityType` 中传递的值不是 `EntityType` 枚举的有效成员，该方法返回 null。</span><span class="sxs-lookup"><span data-stu-id="72c29-868">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="72c29-869">如果没有指定类型的实体的存在于项目的正文中，该方法将返回空数组。</span><span class="sxs-lookup"><span data-stu-id="72c29-869">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="72c29-870">否则，返回的数组中对象的类型取决于 `entityType` 参数中请求实体的类型。</span><span class="sxs-lookup"><span data-stu-id="72c29-870">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="72c29-871">当使用此方法的最低权限级别**受限**时，某些实体类型需要 **ReadItem** 才能进行访问，如下表中所指定。</span><span class="sxs-lookup"><span data-stu-id="72c29-871">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="72c29-872">`entityType` 的值</span><span class="sxs-lookup"><span data-stu-id="72c29-872">Value of `entityType`</span></span> | <span data-ttu-id="72c29-873">返回的数组中对象的类型</span><span class="sxs-lookup"><span data-stu-id="72c29-873">Type of objects in returned array</span></span> | <span data-ttu-id="72c29-874">所需权限级别</span><span class="sxs-lookup"><span data-stu-id="72c29-874">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="72c29-875">String</span><span class="sxs-lookup"><span data-stu-id="72c29-875">String</span></span> | <span data-ttu-id="72c29-876">**受限**</span><span class="sxs-lookup"><span data-stu-id="72c29-876">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="72c29-877">Contact</span><span class="sxs-lookup"><span data-stu-id="72c29-877">Contact</span></span> | <span data-ttu-id="72c29-878">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="72c29-878">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="72c29-879">String</span><span class="sxs-lookup"><span data-stu-id="72c29-879">String</span></span> | <span data-ttu-id="72c29-880">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="72c29-880">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="72c29-881">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="72c29-881">MeetingSuggestion</span></span> | <span data-ttu-id="72c29-882">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="72c29-882">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="72c29-883">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="72c29-883">PhoneNumber</span></span> | <span data-ttu-id="72c29-884">**受限**</span><span class="sxs-lookup"><span data-stu-id="72c29-884">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="72c29-885">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="72c29-885">TaskSuggestion</span></span> | <span data-ttu-id="72c29-886">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="72c29-886">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="72c29-887">String</span><span class="sxs-lookup"><span data-stu-id="72c29-887">String</span></span> | <span data-ttu-id="72c29-888">**受限**</span><span class="sxs-lookup"><span data-stu-id="72c29-888">**Restricted**</span></span> |

<span data-ttu-id="72c29-889">类型：Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="72c29-889">Type: Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="72c29-890">示例</span><span class="sxs-lookup"><span data-stu-id="72c29-890">Example</span></span>

<span data-ttu-id="72c29-891">下面的示例演示如何访问表示当前项目的正文中的邮政地址的字符串数组。</span><span class="sxs-lookup"><span data-stu-id="72c29-891">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion"></a><span data-ttu-id="72c29-892">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="72c29-892">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span></span>

<span data-ttu-id="72c29-893">返回传递清单 XML 文件中定义的命名筛选器的所选项目中的已知实体。</span><span class="sxs-lookup"><span data-stu-id="72c29-893">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="72c29-894">在适用于 iOS 的 Outlook 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="72c29-894">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="72c29-895">`getFilteredEntitiesByName` 方法返回匹配在具有指定 `FilterName` 元素值的清单 XML 文件中的 [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) 规则元素中定义的正则表达式的实体。</span><span class="sxs-lookup"><span data-stu-id="72c29-895">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="72c29-896">参数：</span><span class="sxs-lookup"><span data-stu-id="72c29-896">Parameters:</span></span>

|<span data-ttu-id="72c29-897">名称</span><span class="sxs-lookup"><span data-stu-id="72c29-897">Name</span></span>| <span data-ttu-id="72c29-898">类型</span><span class="sxs-lookup"><span data-stu-id="72c29-898">Type</span></span>| <span data-ttu-id="72c29-899">说明</span><span class="sxs-lookup"><span data-stu-id="72c29-899">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="72c29-900">字符串</span><span class="sxs-lookup"><span data-stu-id="72c29-900">String</span></span>|<span data-ttu-id="72c29-901">定义筛选器匹配的 `ItemHasKnownEntity` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="72c29-901">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="72c29-902">要求</span><span class="sxs-lookup"><span data-stu-id="72c29-902">Requirements</span></span>

|<span data-ttu-id="72c29-903">要求</span><span class="sxs-lookup"><span data-stu-id="72c29-903">Requirement</span></span>| <span data-ttu-id="72c29-904">值</span><span class="sxs-lookup"><span data-stu-id="72c29-904">Value</span></span>|
|---|---|
|[<span data-ttu-id="72c29-905">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="72c29-905">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="72c29-906">1.0</span><span class="sxs-lookup"><span data-stu-id="72c29-906">1.0</span></span>|
|[<span data-ttu-id="72c29-907">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="72c29-907">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="72c29-908">ReadItem</span><span class="sxs-lookup"><span data-stu-id="72c29-908">ReadItem</span></span>|
|[<span data-ttu-id="72c29-909">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="72c29-909">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="72c29-910">阅读</span><span class="sxs-lookup"><span data-stu-id="72c29-910">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="72c29-911">返回：</span><span class="sxs-lookup"><span data-stu-id="72c29-911">Returns:</span></span>

<span data-ttu-id="72c29-p155">如果具有匹配 `name` 参数的 `FilterName` 元素值的清单中没有任何 `ItemHasKnownEntity` 元素，则该方法返回 `null`。如果 `name` 参数匹配清单中的 `ItemHasKnownEntity` 元素，但在匹配的当前项目中没有实体，则该方法返回一个空数组。</span><span class="sxs-lookup"><span data-stu-id="72c29-p155">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="72c29-914">类型：Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="72c29-914">Type: Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="72c29-915">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="72c29-915">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="72c29-916">返回所选项目中匹配在清单 XML 文件中定义的正则表达式的字符串值。</span><span class="sxs-lookup"><span data-stu-id="72c29-916">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="72c29-917">在适用于 iOS 的 Outlook 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="72c29-917">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="72c29-p156">`getRegExMatches` 方法返回匹配在清单 XML 文件中的每个 `ItemHasRegularExpressionMatch` 或 `ItemHasKnownEntity` 规则元素中定义的正则表达式的字符串。对于 `ItemHasRegularExpressionMatch` 规则，匹配字符串必须发生在该规则指定的项目的属性中。`PropertyName` 简单类型定义支持的属性。</span><span class="sxs-lookup"><span data-stu-id="72c29-p156">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="72c29-921">例如，考虑一个外接程序清单具有以下 `Rule` 元素：</span><span class="sxs-lookup"><span data-stu-id="72c29-921">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="72c29-922">从 `getRegExMatches` 返回的对象应有两个属性：`fruits` 和 `veggies`。</span><span class="sxs-lookup"><span data-stu-id="72c29-922">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="72c29-p157">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。而是使用 [`Body.getAsync`](/javascript/api/outlook_1_5/office.body#getasync-coerciontype--options--callback-) 方法检索整个正文。</span><span class="sxs-lookup"><span data-stu-id="72c29-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_5/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="72c29-926">Requirements</span><span class="sxs-lookup"><span data-stu-id="72c29-926">Requirements</span></span>

|<span data-ttu-id="72c29-927">要求</span><span class="sxs-lookup"><span data-stu-id="72c29-927">Requirement</span></span>| <span data-ttu-id="72c29-928">值</span><span class="sxs-lookup"><span data-stu-id="72c29-928">Value</span></span>|
|---|---|
|[<span data-ttu-id="72c29-929">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="72c29-929">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="72c29-930">1.0</span><span class="sxs-lookup"><span data-stu-id="72c29-930">1.0</span></span>|
|[<span data-ttu-id="72c29-931">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="72c29-931">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="72c29-932">ReadItem</span><span class="sxs-lookup"><span data-stu-id="72c29-932">ReadItem</span></span>|
|[<span data-ttu-id="72c29-933">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="72c29-933">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="72c29-934">阅读</span><span class="sxs-lookup"><span data-stu-id="72c29-934">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="72c29-935">返回：</span><span class="sxs-lookup"><span data-stu-id="72c29-935">Returns:</span></span>

<span data-ttu-id="72c29-p158">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串数组的对象。每个数组的名称等于匹配 `ItemHasRegularExpressionMatch` 规则的 `RegExName` 属性或匹配 `ItemHasKnownEntity` 规则的 `FilterName` 属性的相应值。</span><span class="sxs-lookup"><span data-stu-id="72c29-p158">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="72c29-938">

<dt>
类型</dt>


</span><span class="sxs-lookup"><span data-stu-id="72c29-938">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="72c29-939">对象</span><span class="sxs-lookup"><span data-stu-id="72c29-939">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="72c29-940">示例</span><span class="sxs-lookup"><span data-stu-id="72c29-940">Example</span></span>

<span data-ttu-id="72c29-941">以下示例显示了如何访问正则表达式 <rule> 元素 `fruits` 和 `veggies` 的匹配项的数组，这些元素在清单中指定。</rule></span><span class="sxs-lookup"><span data-stu-id="72c29-941">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="72c29-942">getregexmatchesbyname （name） → (nullable) {数组。 < 字符串 >}</span><span class="sxs-lookup"><span data-stu-id="72c29-942">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="72c29-943">返回匹配在清单 XML 文件中定义的命名正则表达式的所选项目中的字符串值。</span><span class="sxs-lookup"><span data-stu-id="72c29-943">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="72c29-944">在适用于 iOS 的 Outlook 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="72c29-944">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="72c29-945">`getRegExMatchesByName` 方法返回匹配在具有指定 `RegExName` 元素值的清单 XML 文件中的 `ItemHasRegularExpressionMatch` 规则元素中定义的正则表达式的字符串。</span><span class="sxs-lookup"><span data-stu-id="72c29-945">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="72c29-p159">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。</span><span class="sxs-lookup"><span data-stu-id="72c29-p159">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="72c29-948">参数：</span><span class="sxs-lookup"><span data-stu-id="72c29-948">Parameters:</span></span>

|<span data-ttu-id="72c29-949">名称</span><span class="sxs-lookup"><span data-stu-id="72c29-949">Name</span></span>| <span data-ttu-id="72c29-950">类型</span><span class="sxs-lookup"><span data-stu-id="72c29-950">Type</span></span>| <span data-ttu-id="72c29-951">说明</span><span class="sxs-lookup"><span data-stu-id="72c29-951">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="72c29-952">字符串</span><span class="sxs-lookup"><span data-stu-id="72c29-952">String</span></span>|<span data-ttu-id="72c29-953">定义筛选器匹配的 `ItemHasRegularExpressionMatch` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="72c29-953">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="72c29-954">要求</span><span class="sxs-lookup"><span data-stu-id="72c29-954">Requirements</span></span>

|<span data-ttu-id="72c29-955">要求</span><span class="sxs-lookup"><span data-stu-id="72c29-955">Requirement</span></span>| <span data-ttu-id="72c29-956">值</span><span class="sxs-lookup"><span data-stu-id="72c29-956">Value</span></span>|
|---|---|
|[<span data-ttu-id="72c29-957">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="72c29-957">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="72c29-958">1.0</span><span class="sxs-lookup"><span data-stu-id="72c29-958">1.0</span></span>|
|[<span data-ttu-id="72c29-959">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="72c29-959">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="72c29-960">ReadItem</span><span class="sxs-lookup"><span data-stu-id="72c29-960">ReadItem</span></span>|
|[<span data-ttu-id="72c29-961">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="72c29-961">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="72c29-962">阅读</span><span class="sxs-lookup"><span data-stu-id="72c29-962">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="72c29-963">返回：</span><span class="sxs-lookup"><span data-stu-id="72c29-963">Returns:</span></span>

<span data-ttu-id="72c29-964">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串的数组。</span><span class="sxs-lookup"><span data-stu-id="72c29-964">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="72c29-965">

<dt>类型</dt>

</span><span class="sxs-lookup"><span data-stu-id="72c29-965">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="72c29-966">数组。 < 字符串 ></span><span class="sxs-lookup"><span data-stu-id="72c29-966">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="72c29-967">示例</span><span class="sxs-lookup"><span data-stu-id="72c29-967">Example</span></span>

```
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="72c29-968">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="72c29-968">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="72c29-969">以异步方式返回邮件的主题或正文中选定的数据。</span><span class="sxs-lookup"><span data-stu-id="72c29-969">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="72c29-p160">如果没有选定内容，但光标位于正文或主题中，此方法将会为所选数据返回 null。如果选定的是字段，而不是正文或主题，则此方法返回 `InvalidSelection` 错误。</span><span class="sxs-lookup"><span data-stu-id="72c29-p160">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="72c29-972">参数：</span><span class="sxs-lookup"><span data-stu-id="72c29-972">Parameters:</span></span>

|<span data-ttu-id="72c29-973">名称</span><span class="sxs-lookup"><span data-stu-id="72c29-973">Name</span></span>| <span data-ttu-id="72c29-974">类型</span><span class="sxs-lookup"><span data-stu-id="72c29-974">Type</span></span>| <span data-ttu-id="72c29-975">属性</span><span class="sxs-lookup"><span data-stu-id="72c29-975">Attributes</span></span>| <span data-ttu-id="72c29-976">说明</span><span class="sxs-lookup"><span data-stu-id="72c29-976">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="72c29-977">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="72c29-977">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="72c29-p161">请求数据的格式。如果为文本，则此方法返回纯文本作为字符串，删除任何显示的 HTML 标记。如果为 HTML，则此方法返回所选文本，不论是纯文本还是 HTML。</span><span class="sxs-lookup"><span data-stu-id="72c29-p161">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="72c29-981">对象</span><span class="sxs-lookup"><span data-stu-id="72c29-981">Object</span></span>| <span data-ttu-id="72c29-982">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="72c29-982">&lt;optional&gt;</span></span>|<span data-ttu-id="72c29-983">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="72c29-983">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="72c29-984">对象</span><span class="sxs-lookup"><span data-stu-id="72c29-984">Object</span></span>| <span data-ttu-id="72c29-985">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="72c29-985">&lt;optional&gt;</span></span>|<span data-ttu-id="72c29-986">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="72c29-986">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="72c29-987">函数</span><span class="sxs-lookup"><span data-stu-id="72c29-987">function</span></span>||<span data-ttu-id="72c29-988">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="72c29-988">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="72c29-989">若要从回调方法访问所选数据，请调用 `asyncResult.value.data`。</span><span class="sxs-lookup"><span data-stu-id="72c29-989">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="72c29-990">若要访问所选内容来自源属性，请调用`asyncResult.value.sourceProperty`，其中将为`body`或`subject`。</span><span class="sxs-lookup"><span data-stu-id="72c29-990">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="72c29-991">要求</span><span class="sxs-lookup"><span data-stu-id="72c29-991">Requirements</span></span>

|<span data-ttu-id="72c29-992">要求</span><span class="sxs-lookup"><span data-stu-id="72c29-992">Requirement</span></span>| <span data-ttu-id="72c29-993">值</span><span class="sxs-lookup"><span data-stu-id="72c29-993">Value</span></span>|
|---|---|
|[<span data-ttu-id="72c29-994">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="72c29-994">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="72c29-995">1.2</span><span class="sxs-lookup"><span data-stu-id="72c29-995">1.2</span></span>|
|[<span data-ttu-id="72c29-996">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="72c29-996">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="72c29-997">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="72c29-997">ReadWriteItem</span></span>|
|[<span data-ttu-id="72c29-998">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="72c29-998">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="72c29-999">撰写</span><span class="sxs-lookup"><span data-stu-id="72c29-999">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="72c29-1000">返回：</span><span class="sxs-lookup"><span data-stu-id="72c29-1000">Returns:</span></span>

<span data-ttu-id="72c29-1001">作为字符串的所选数据的格式由 `coercionType` 确定。</span><span class="sxs-lookup"><span data-stu-id="72c29-1001">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="72c29-1002">

<dt>
类型</dt>


</span><span class="sxs-lookup"><span data-stu-id="72c29-1002">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="72c29-1003">字符串</span><span class="sxs-lookup"><span data-stu-id="72c29-1003">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="72c29-1004">示例</span><span class="sxs-lookup"><span data-stu-id="72c29-1004">Example</span></span>

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

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="72c29-1005">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="72c29-1005">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="72c29-1006">异步加载所选项目上此外接程序的自定义属性。</span><span class="sxs-lookup"><span data-stu-id="72c29-1006">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="72c29-p163">自定义属性基于每个应用、每个项目存储为键/值对。此方法在回调中返回 `CustomProperties` 对象，该回调提供访问特定于当前项目和当前外接程序的自定义属性的方法。自定义属性未在项目上加密，因此这不应用作安全存储。</span><span class="sxs-lookup"><span data-stu-id="72c29-p163">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="72c29-1010">参数：</span><span class="sxs-lookup"><span data-stu-id="72c29-1010">Parameters:</span></span>

|<span data-ttu-id="72c29-1011">名称</span><span class="sxs-lookup"><span data-stu-id="72c29-1011">Name</span></span>| <span data-ttu-id="72c29-1012">类型</span><span class="sxs-lookup"><span data-stu-id="72c29-1012">Type</span></span>| <span data-ttu-id="72c29-1013">属性</span><span class="sxs-lookup"><span data-stu-id="72c29-1013">Attributes</span></span>| <span data-ttu-id="72c29-1014">说明</span><span class="sxs-lookup"><span data-stu-id="72c29-1014">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="72c29-1015">函数</span><span class="sxs-lookup"><span data-stu-id="72c29-1015">function</span></span>||<span data-ttu-id="72c29-1016">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="72c29-1016">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="72c29-1017">自定义属性作为 `asyncResult.value` 属性中的 [`CustomProperties`](/javascript/api/outlook_1_5/office.customproperties) 对象提供。</span><span class="sxs-lookup"><span data-stu-id="72c29-1017">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_5/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="72c29-1018">此对象可用于获取、 设置，和从项目中删除自定义属性并将更改保存到自定义属性设置回服务器。</span><span class="sxs-lookup"><span data-stu-id="72c29-1018">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="72c29-1019">对象</span><span class="sxs-lookup"><span data-stu-id="72c29-1019">Object</span></span>| <span data-ttu-id="72c29-1020">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="72c29-1020">&lt;optional&gt;</span></span>|<span data-ttu-id="72c29-1021">开发人员可以提供他们想要访问的回调函数中的任何对象。</span><span class="sxs-lookup"><span data-stu-id="72c29-1021">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="72c29-1022">此对象可以访问`asyncResult.asyncContext`的回调函数中的属性。</span><span class="sxs-lookup"><span data-stu-id="72c29-1022">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="72c29-1023">要求</span><span class="sxs-lookup"><span data-stu-id="72c29-1023">Requirements</span></span>

|<span data-ttu-id="72c29-1024">要求</span><span class="sxs-lookup"><span data-stu-id="72c29-1024">Requirement</span></span>| <span data-ttu-id="72c29-1025">值</span><span class="sxs-lookup"><span data-stu-id="72c29-1025">Value</span></span>|
|---|---|
|[<span data-ttu-id="72c29-1026">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="72c29-1026">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="72c29-1027">1.0</span><span class="sxs-lookup"><span data-stu-id="72c29-1027">1.0</span></span>|
|[<span data-ttu-id="72c29-1028">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="72c29-1028">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="72c29-1029">ReadItem</span><span class="sxs-lookup"><span data-stu-id="72c29-1029">ReadItem</span></span>|
|[<span data-ttu-id="72c29-1030">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="72c29-1030">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="72c29-1031">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="72c29-1031">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="72c29-1032">示例</span><span class="sxs-lookup"><span data-stu-id="72c29-1032">Example</span></span>

<span data-ttu-id="72c29-p166">以下代码示例显示了如何使用 `loadCustomPropertiesAsync` 方法异步加载特定于当前项目的自定义属性。该示例还显示了如何使用 `CustomProperties.saveAsync` 方法将这些属性重新保存到服务器。加载自定义属性后，该代码示例将使用 `CustomProperties.get` 方法读取自定义属性 `myProp`，使用 `CustomProperties.set` 方法写入自定义属性 `otherProp`，最后调用 `saveAsync` 方法保存这些自定义属性。</span><span class="sxs-lookup"><span data-stu-id="72c29-p166">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="72c29-1036">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="72c29-1036">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="72c29-1037">将附件从邮件或约会中删除。</span><span class="sxs-lookup"><span data-stu-id="72c29-1037">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="72c29-p167">`removeAttachmentAsync` 方法删除项目中带指定标识符的附件。最佳做法是，仅当同一个邮件应用程序在同一会话中添加了一个附件时，你才应使用该附件标识符来删除该附件。在 Outlook Web App 和适用于设备的 OWA 中，附件标识符只在同一个会话中才有效。当用户关闭应用程序，或者如果用户开始在内嵌窗体中撰写，并在随后弹出的内嵌窗体中继续在单独的窗口撰写时，会话即结束。</span><span class="sxs-lookup"><span data-stu-id="72c29-p167">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="72c29-1042">参数：</span><span class="sxs-lookup"><span data-stu-id="72c29-1042">Parameters:</span></span>

|<span data-ttu-id="72c29-1043">名称</span><span class="sxs-lookup"><span data-stu-id="72c29-1043">Name</span></span>| <span data-ttu-id="72c29-1044">类型</span><span class="sxs-lookup"><span data-stu-id="72c29-1044">Type</span></span>| <span data-ttu-id="72c29-1045">属性</span><span class="sxs-lookup"><span data-stu-id="72c29-1045">Attributes</span></span>| <span data-ttu-id="72c29-1046">说明</span><span class="sxs-lookup"><span data-stu-id="72c29-1046">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="72c29-1047">字符串</span><span class="sxs-lookup"><span data-stu-id="72c29-1047">String</span></span>||<span data-ttu-id="72c29-p168">要删除的附件的标识符。字符串的最大长度为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="72c29-p168">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`| <span data-ttu-id="72c29-1050">对象</span><span class="sxs-lookup"><span data-stu-id="72c29-1050">Object</span></span>| <span data-ttu-id="72c29-1051">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="72c29-1051">&lt;optional&gt;</span></span>|<span data-ttu-id="72c29-1052">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="72c29-1052">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="72c29-1053">对象</span><span class="sxs-lookup"><span data-stu-id="72c29-1053">Object</span></span>| <span data-ttu-id="72c29-1054">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="72c29-1054">&lt;optional&gt;</span></span>|<span data-ttu-id="72c29-1055">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="72c29-1055">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="72c29-1056">函数</span><span class="sxs-lookup"><span data-stu-id="72c29-1056">function</span></span>| <span data-ttu-id="72c29-1057">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="72c29-1057">&lt;optional&gt;</span></span>|<span data-ttu-id="72c29-1058">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="72c29-1058">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="72c29-1059">如果删除附件失败，`asyncResult.error` 属性将包含一个说明失败原因的错误代码。</span><span class="sxs-lookup"><span data-stu-id="72c29-1059">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="72c29-1060">错误</span><span class="sxs-lookup"><span data-stu-id="72c29-1060">Errors</span></span>

| <span data-ttu-id="72c29-1061">错误代码</span><span class="sxs-lookup"><span data-stu-id="72c29-1061">Error code</span></span> | <span data-ttu-id="72c29-1062">说明</span><span class="sxs-lookup"><span data-stu-id="72c29-1062">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="72c29-1063">附件标识符不存在。</span><span class="sxs-lookup"><span data-stu-id="72c29-1063">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="72c29-1064">要求</span><span class="sxs-lookup"><span data-stu-id="72c29-1064">Requirements</span></span>

|<span data-ttu-id="72c29-1065">要求</span><span class="sxs-lookup"><span data-stu-id="72c29-1065">Requirement</span></span>| <span data-ttu-id="72c29-1066">值</span><span class="sxs-lookup"><span data-stu-id="72c29-1066">Value</span></span>|
|---|---|
|[<span data-ttu-id="72c29-1067">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="72c29-1067">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="72c29-1068">1.1</span><span class="sxs-lookup"><span data-stu-id="72c29-1068">1.1</span></span>|
|[<span data-ttu-id="72c29-1069">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="72c29-1069">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="72c29-1070">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="72c29-1070">ReadWriteItem</span></span>|
|[<span data-ttu-id="72c29-1071">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="72c29-1071">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="72c29-1072">撰写</span><span class="sxs-lookup"><span data-stu-id="72c29-1072">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="72c29-1073">示例</span><span class="sxs-lookup"><span data-stu-id="72c29-1073">Example</span></span>

<span data-ttu-id="72c29-1074">以下代码删除包含标识符 '0' 的附件。</span><span class="sxs-lookup"><span data-stu-id="72c29-1074">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="72c29-1075">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="72c29-1075">saveAsync([options], callback)</span></span>

<span data-ttu-id="72c29-1076">异步保存项目。</span><span class="sxs-lookup"><span data-stu-id="72c29-1076">Asynchronously saves an item.</span></span>

<span data-ttu-id="72c29-p169">调用时，此方法将当前邮件保存为草稿，并通过回调方法返回项目 ID。在 Outlook Web App 或 Outlook 联机模式下，该项目被保存到服务器中。在 Outlook 缓存模式下，该项目被保存到本地缓存中。</span><span class="sxs-lookup"><span data-stu-id="72c29-p169">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="72c29-1080">如果加载项调用`saveAsync`中的项目在撰写模式下才能获取`itemId`若要使用 EWS 或 REST API，请注意，缓存模式 Outlook 时，可能需要一些时间才能项目实际同步到服务器。</span><span class="sxs-lookup"><span data-stu-id="72c29-1080">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="72c29-1081">直到该项目同步，使用`itemId`将返回错误。</span><span class="sxs-lookup"><span data-stu-id="72c29-1081">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="72c29-p171">由于约会没有草稿状态，如果以撰写模式在约会中调用 `saveAsync`，则该项将被保存为用户日历中的正常约会。对于之前未保存过的新约会，则不会发送邀请。保存现有约会将向添加或删除的与会者发送更新。</span><span class="sxs-lookup"><span data-stu-id="72c29-p171">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="72c29-1085">以下客户端具有不同行为的`saveAsync`上约会的撰写模式下：</span><span class="sxs-lookup"><span data-stu-id="72c29-1085">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="72c29-1086">Mac Outlook 不支持`saveAsync`中的会议在撰写模式。</span><span class="sxs-lookup"><span data-stu-id="72c29-1086">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="72c29-1087">调用`saveAsync`上 Mac Outlook 中的会议，则将返回错误。</span><span class="sxs-lookup"><span data-stu-id="72c29-1087">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="72c29-1088">在 web 上的 outlook 始终发送的邀请或更新何时`saveAsync`上约会调用在撰写模式。</span><span class="sxs-lookup"><span data-stu-id="72c29-1088">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="72c29-1089">参数：</span><span class="sxs-lookup"><span data-stu-id="72c29-1089">Parameters:</span></span>

|<span data-ttu-id="72c29-1090">名称</span><span class="sxs-lookup"><span data-stu-id="72c29-1090">Name</span></span>| <span data-ttu-id="72c29-1091">类型</span><span class="sxs-lookup"><span data-stu-id="72c29-1091">Type</span></span>| <span data-ttu-id="72c29-1092">属性</span><span class="sxs-lookup"><span data-stu-id="72c29-1092">Attributes</span></span>| <span data-ttu-id="72c29-1093">说明</span><span class="sxs-lookup"><span data-stu-id="72c29-1093">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="72c29-1094">对象</span><span class="sxs-lookup"><span data-stu-id="72c29-1094">Object</span></span>| <span data-ttu-id="72c29-1095">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="72c29-1095">&lt;optional&gt;</span></span>|<span data-ttu-id="72c29-1096">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="72c29-1096">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="72c29-1097">对象</span><span class="sxs-lookup"><span data-stu-id="72c29-1097">Object</span></span>| <span data-ttu-id="72c29-1098">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="72c29-1098">&lt;optional&gt;</span></span>|<span data-ttu-id="72c29-1099">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="72c29-1099">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="72c29-1100">函数</span><span class="sxs-lookup"><span data-stu-id="72c29-1100">function</span></span>||<span data-ttu-id="72c29-1101">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="72c29-1101">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="72c29-1102">中提供的项标识符是成功，`asyncResult.value`属性。</span><span class="sxs-lookup"><span data-stu-id="72c29-1102">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="72c29-1103">要求</span><span class="sxs-lookup"><span data-stu-id="72c29-1103">Requirements</span></span>

|<span data-ttu-id="72c29-1104">要求</span><span class="sxs-lookup"><span data-stu-id="72c29-1104">Requirement</span></span>| <span data-ttu-id="72c29-1105">值</span><span class="sxs-lookup"><span data-stu-id="72c29-1105">Value</span></span>|
|---|---|
|[<span data-ttu-id="72c29-1106">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="72c29-1106">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="72c29-1107">1.3</span><span class="sxs-lookup"><span data-stu-id="72c29-1107">1.3</span></span>|
|[<span data-ttu-id="72c29-1108">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="72c29-1108">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="72c29-1109">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="72c29-1109">ReadWriteItem</span></span>|
|[<span data-ttu-id="72c29-1110">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="72c29-1110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="72c29-1111">撰写</span><span class="sxs-lookup"><span data-stu-id="72c29-1111">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="72c29-1112">示例</span><span class="sxs-lookup"><span data-stu-id="72c29-1112">Examples</span></span>

```
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="72c29-p173">下面是传递给回调函数的 `result` 参数的示例。`value` 属性包含的项目的项目 ID。</span><span class="sxs-lookup"><span data-stu-id="72c29-p173">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="72c29-1115">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="72c29-1115">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="72c29-1116">以异步方式将数据插入到邮件的正文或主题中。</span><span class="sxs-lookup"><span data-stu-id="72c29-1116">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="72c29-p174">`setSelectedDataAsync` 方法将指定字符串插入到项目主题或正文的光标位置，或者，如果在编辑器中已选择文本，则该方法将替换选择的文本。如果光标不在正文或主题字段中，则返回错误。插入之后，光标会位于插入内容的末尾。</span><span class="sxs-lookup"><span data-stu-id="72c29-p174">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="72c29-1120">参数：</span><span class="sxs-lookup"><span data-stu-id="72c29-1120">Parameters:</span></span>

|<span data-ttu-id="72c29-1121">名称</span><span class="sxs-lookup"><span data-stu-id="72c29-1121">Name</span></span>| <span data-ttu-id="72c29-1122">类型</span><span class="sxs-lookup"><span data-stu-id="72c29-1122">Type</span></span>| <span data-ttu-id="72c29-1123">属性</span><span class="sxs-lookup"><span data-stu-id="72c29-1123">Attributes</span></span>| <span data-ttu-id="72c29-1124">说明</span><span class="sxs-lookup"><span data-stu-id="72c29-1124">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="72c29-1125">字符串</span><span class="sxs-lookup"><span data-stu-id="72c29-1125">String</span></span>||<span data-ttu-id="72c29-p175">要插入的数据。数据不得超过 1,000,000 个字符。如果传入的数据超过 1,000,000 个字符，则会引发 `ArgumentOutOfRange` 异常。</span><span class="sxs-lookup"><span data-stu-id="72c29-p175">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="72c29-1129">Object</span><span class="sxs-lookup"><span data-stu-id="72c29-1129">Object</span></span>| <span data-ttu-id="72c29-1130">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="72c29-1130">&lt;optional&gt;</span></span>|<span data-ttu-id="72c29-1131">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="72c29-1131">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="72c29-1132">对象</span><span class="sxs-lookup"><span data-stu-id="72c29-1132">Object</span></span>| <span data-ttu-id="72c29-1133">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="72c29-1133">&lt;optional&gt;</span></span>|<span data-ttu-id="72c29-1134">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="72c29-1134">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="72c29-1135">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="72c29-1135">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="72c29-1136">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="72c29-1136">&lt;optional&gt;</span></span>|<span data-ttu-id="72c29-p176">如果为 `text`，则在 Outlook Web App 和 Outlook 中应用当前样式。如果该字段是 HTML 编辑器，则仅插入文本数据，即使数据为 HTML。</span><span class="sxs-lookup"><span data-stu-id="72c29-p176">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="72c29-p177">如果 `html` 和该字段支持 HTML（主题不支持），则在 Outlook Web App 中应用当前样式，而在 Outlook 中应用默认样式。如果该字段是文本字段，则返回 `InvalidDataFormat` 错误。</span><span class="sxs-lookup"><span data-stu-id="72c29-p177">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="72c29-1141">如果未设置 `coercionType`，则结果取决于该字段：如果该字段是 HTML，则使用 HTML；如果该字段是文本，则使用纯文本。</span><span class="sxs-lookup"><span data-stu-id="72c29-1141">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="72c29-1142">function</span><span class="sxs-lookup"><span data-stu-id="72c29-1142">function</span></span>||<span data-ttu-id="72c29-1143">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="72c29-1143">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="72c29-1144">Requirements</span><span class="sxs-lookup"><span data-stu-id="72c29-1144">Requirements</span></span>

|<span data-ttu-id="72c29-1145">要求</span><span class="sxs-lookup"><span data-stu-id="72c29-1145">Requirement</span></span>| <span data-ttu-id="72c29-1146">值</span><span class="sxs-lookup"><span data-stu-id="72c29-1146">Value</span></span>|
|---|---|
|[<span data-ttu-id="72c29-1147">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="72c29-1147">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="72c29-1148">1.2</span><span class="sxs-lookup"><span data-stu-id="72c29-1148">1.2</span></span>|
|[<span data-ttu-id="72c29-1149">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="72c29-1149">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="72c29-1150">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="72c29-1150">ReadWriteItem</span></span>|
|[<span data-ttu-id="72c29-1151">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="72c29-1151">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="72c29-1152">撰写</span><span class="sxs-lookup"><span data-stu-id="72c29-1152">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="72c29-1153">示例</span><span class="sxs-lookup"><span data-stu-id="72c29-1153">Example</span></span>

```
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```