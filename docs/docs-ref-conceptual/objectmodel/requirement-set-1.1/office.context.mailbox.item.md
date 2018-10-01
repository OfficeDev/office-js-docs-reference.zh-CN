
# <a name="item"></a><span data-ttu-id="a1c20-101">item</span><span class="sxs-lookup"><span data-stu-id="a1c20-101">item</span></span>

### <span data-ttu-id="a1c20-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span><span class="sxs-lookup"><span data-stu-id="a1c20-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span></span>

<span data-ttu-id="a1c20-p102">`item` 命名空间用于访问当前选定的邮件、会议请求或约会。可以通过使用 [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook11officemailboxenumsitemtype) 属性确定 `item` 的类型。</span><span class="sxs-lookup"><span data-stu-id="a1c20-p102">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook11officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="a1c20-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="a1c20-106">Requirements</span></span>

|<span data-ttu-id="a1c20-107">要求</span><span class="sxs-lookup"><span data-stu-id="a1c20-107">Requirement</span></span>| <span data-ttu-id="a1c20-108">值</span><span class="sxs-lookup"><span data-stu-id="a1c20-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="a1c20-109">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="a1c20-109">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a1c20-110">1.0</span><span class="sxs-lookup"><span data-stu-id="a1c20-110">1.0</span></span>|
|[<span data-ttu-id="a1c20-111">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="a1c20-111">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a1c20-112">受限</span><span class="sxs-lookup"><span data-stu-id="a1c20-112">Restricted</span></span>|
|[<span data-ttu-id="a1c20-113">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="a1c20-113">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a1c20-114">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="a1c20-114">Compose or read</span></span>|

### <a name="example"></a><span data-ttu-id="a1c20-115">示例</span><span class="sxs-lookup"><span data-stu-id="a1c20-115">Example</span></span>

<span data-ttu-id="a1c20-116">以下 JavaScript 代码示例显示了如何访问 Outlook 中当前项目的 `subject` 属性。</span><span class="sxs-lookup"><span data-stu-id="a1c20-116">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```JavaScript
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

### <a name="members"></a><span data-ttu-id="a1c20-117">成员</span><span class="sxs-lookup"><span data-stu-id="a1c20-117">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook11officeattachmentdetails"></a><span data-ttu-id="a1c20-118">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="a1c20-118">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)></span></span>

<span data-ttu-id="a1c20-p103">获取项目的附件数组。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="a1c20-p103">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="a1c20-121">某些类型的文件被阻止 Outlook 由于潜在的安全问题，并因此不返回。</span><span class="sxs-lookup"><span data-stu-id="a1c20-121">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="a1c20-122">有关详细信息，请参阅[在 Outlook 中被阻止的附件](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)。</span><span class="sxs-lookup"><span data-stu-id="a1c20-122">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="a1c20-123">类型:</span><span class="sxs-lookup"><span data-stu-id="a1c20-123">Type:</span></span>

*   <span data-ttu-id="a1c20-124">Array.<[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="a1c20-124">Array.<[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="a1c20-125">要求</span><span class="sxs-lookup"><span data-stu-id="a1c20-125">Requirements</span></span>

|<span data-ttu-id="a1c20-126">要求</span><span class="sxs-lookup"><span data-stu-id="a1c20-126">Requirement</span></span>| <span data-ttu-id="a1c20-127">值</span><span class="sxs-lookup"><span data-stu-id="a1c20-127">Value</span></span>|
|---|---|
|[<span data-ttu-id="a1c20-128">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="a1c20-128">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a1c20-129">1.0</span><span class="sxs-lookup"><span data-stu-id="a1c20-129">1.0</span></span>|
|[<span data-ttu-id="a1c20-130">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="a1c20-130">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a1c20-131">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a1c20-131">ReadItem</span></span>|
|[<span data-ttu-id="a1c20-132">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="a1c20-132">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a1c20-133">阅读</span><span class="sxs-lookup"><span data-stu-id="a1c20-133">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a1c20-134">示例</span><span class="sxs-lookup"><span data-stu-id="a1c20-134">Example</span></span>

<span data-ttu-id="a1c20-135">以下代码使用当前项目上所有附件的详细信息构成 HTML 字符串。</span><span class="sxs-lookup"><span data-stu-id="a1c20-135">The following code builds an HTML string with details of all attachments on the current item.</span></span>

```JavaScript
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

####  <a name="bcc-recipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="a1c20-136">bcc :[Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a1c20-136">bcc :[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="a1c20-137">获取一个对象，提供用于获取或更新邮件的密件抄送 （密件抄送副本） 行上的收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="a1c20-137">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="a1c20-138">仅限撰写模式。</span><span class="sxs-lookup"><span data-stu-id="a1c20-138">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="a1c20-139">类型：</span><span class="sxs-lookup"><span data-stu-id="a1c20-139">Type:</span></span>

*   [<span data-ttu-id="a1c20-140">收件人</span><span class="sxs-lookup"><span data-stu-id="a1c20-140">Recipients</span></span>](/javascript/api/outlook_1_1/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="a1c20-141">要求</span><span class="sxs-lookup"><span data-stu-id="a1c20-141">Requirements</span></span>

|<span data-ttu-id="a1c20-142">要求</span><span class="sxs-lookup"><span data-stu-id="a1c20-142">Requirement</span></span>| <span data-ttu-id="a1c20-143">值</span><span class="sxs-lookup"><span data-stu-id="a1c20-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="a1c20-144">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="a1c20-144">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a1c20-145">1.1</span><span class="sxs-lookup"><span data-stu-id="a1c20-145">1.1</span></span>|
|[<span data-ttu-id="a1c20-146">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="a1c20-146">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a1c20-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a1c20-147">ReadItem</span></span>|
|[<span data-ttu-id="a1c20-148">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="a1c20-148">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a1c20-149">撰写</span><span class="sxs-lookup"><span data-stu-id="a1c20-149">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="a1c20-150">示例</span><span class="sxs-lookup"><span data-stu-id="a1c20-150">Example</span></span>

```JavaScript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook11officebody"></a><span data-ttu-id="a1c20-151">body :[Body](/javascript/api/outlook_1_1/office.body)</span><span class="sxs-lookup"><span data-stu-id="a1c20-151">body :[Body](/javascript/api/outlook_1_1/office.body)</span></span>

<span data-ttu-id="a1c20-152">获取一个提供用于处理项目正文的方法的对象。</span><span class="sxs-lookup"><span data-stu-id="a1c20-152">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="a1c20-153">类型：</span><span class="sxs-lookup"><span data-stu-id="a1c20-153">Type:</span></span>

*   [<span data-ttu-id="a1c20-154">Body</span><span class="sxs-lookup"><span data-stu-id="a1c20-154">Body</span></span>](/javascript/api/outlook_1_1/office.body)

##### <a name="requirements"></a><span data-ttu-id="a1c20-155">要求</span><span class="sxs-lookup"><span data-stu-id="a1c20-155">Requirements</span></span>

|<span data-ttu-id="a1c20-156">要求</span><span class="sxs-lookup"><span data-stu-id="a1c20-156">Requirement</span></span>| <span data-ttu-id="a1c20-157">值</span><span class="sxs-lookup"><span data-stu-id="a1c20-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="a1c20-158">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="a1c20-158">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a1c20-159">1.1</span><span class="sxs-lookup"><span data-stu-id="a1c20-159">1.1</span></span>|
|[<span data-ttu-id="a1c20-160">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="a1c20-160">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a1c20-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a1c20-161">ReadItem</span></span>|
|[<span data-ttu-id="a1c20-162">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="a1c20-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a1c20-163">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="a1c20-163">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="a1c20-164">抄送： 数组。 <[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[收件人](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a1c20-164">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="a1c20-165">提供对邮件的抄送 (cc) 收件人访问。</span><span class="sxs-lookup"><span data-stu-id="a1c20-165">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="a1c20-166">对象的类型和访问级别取决于当前项的模式。</span><span class="sxs-lookup"><span data-stu-id="a1c20-166">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a1c20-167">阅读模式</span><span class="sxs-lookup"><span data-stu-id="a1c20-167">Read mode</span></span>

<span data-ttu-id="a1c20-p107">`cc` 属性返回包含邮件的**抄送**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="a1c20-p107">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="a1c20-170">撰写模式</span><span class="sxs-lookup"><span data-stu-id="a1c20-170">Compose mode</span></span>

<span data-ttu-id="a1c20-171">`cc`属性返回`Recipients`提供方法用于获取或更新的邮件**Cc**行上的收件人对象。</span><span class="sxs-lookup"><span data-stu-id="a1c20-171">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="a1c20-172">类型：</span><span class="sxs-lookup"><span data-stu-id="a1c20-172">Type:</span></span>

*   <span data-ttu-id="a1c20-173">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a1c20-173">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a1c20-174">要求</span><span class="sxs-lookup"><span data-stu-id="a1c20-174">Requirements</span></span>

|<span data-ttu-id="a1c20-175">要求</span><span class="sxs-lookup"><span data-stu-id="a1c20-175">Requirement</span></span>| <span data-ttu-id="a1c20-176">值</span><span class="sxs-lookup"><span data-stu-id="a1c20-176">Value</span></span>|
|---|---|
|[<span data-ttu-id="a1c20-177">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="a1c20-177">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a1c20-178">1.0</span><span class="sxs-lookup"><span data-stu-id="a1c20-178">1.0</span></span>|
|[<span data-ttu-id="a1c20-179">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="a1c20-179">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a1c20-180">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a1c20-180">ReadItem</span></span>|
|[<span data-ttu-id="a1c20-181">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="a1c20-181">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a1c20-182">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="a1c20-182">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a1c20-183">示例</span><span class="sxs-lookup"><span data-stu-id="a1c20-183">Example</span></span>

```JavaScript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="a1c20-184">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="a1c20-184">(nullable) conversationId :String</span></span>

<span data-ttu-id="a1c20-185">获取包含特定消息的电子邮件会话的标识符。</span><span class="sxs-lookup"><span data-stu-id="a1c20-185">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="a1c20-p108">如果在阅读窗体或撰写窗体的回复中激活邮件应用程序，则此属性可以获得一个整数值。如果用户随后更改了回复邮件的主题（若发送回复），则该邮件的对话 ID 将改变且之前获取的值将不适用。</span><span class="sxs-lookup"><span data-stu-id="a1c20-p108">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="a1c20-p109">对于撰写窗体的新项目，此属性获得一个 null 值。如果用户设置一个主题并保存该项目，`conversationId` 属性将返回一个值。</span><span class="sxs-lookup"><span data-stu-id="a1c20-p109">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="a1c20-190">类型:</span><span class="sxs-lookup"><span data-stu-id="a1c20-190">Type:</span></span>

*   <span data-ttu-id="a1c20-191">String</span><span class="sxs-lookup"><span data-stu-id="a1c20-191">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a1c20-192">要求</span><span class="sxs-lookup"><span data-stu-id="a1c20-192">Requirements</span></span>

|<span data-ttu-id="a1c20-193">要求</span><span class="sxs-lookup"><span data-stu-id="a1c20-193">Requirement</span></span>| <span data-ttu-id="a1c20-194">值</span><span class="sxs-lookup"><span data-stu-id="a1c20-194">Value</span></span>|
|---|---|
|[<span data-ttu-id="a1c20-195">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="a1c20-195">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a1c20-196">1.0</span><span class="sxs-lookup"><span data-stu-id="a1c20-196">1.0</span></span>|
|[<span data-ttu-id="a1c20-197">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="a1c20-197">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a1c20-198">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a1c20-198">ReadItem</span></span>|
|[<span data-ttu-id="a1c20-199">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="a1c20-199">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a1c20-200">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="a1c20-200">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="a1c20-201">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="a1c20-201">dateTimeCreated :Date</span></span>

<span data-ttu-id="a1c20-p110">获取项目创建的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="a1c20-p110">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="a1c20-204">类型：</span><span class="sxs-lookup"><span data-stu-id="a1c20-204">Type:</span></span>

*   <span data-ttu-id="a1c20-205">日期</span><span class="sxs-lookup"><span data-stu-id="a1c20-205">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="a1c20-206">要求</span><span class="sxs-lookup"><span data-stu-id="a1c20-206">Requirements</span></span>

|<span data-ttu-id="a1c20-207">要求</span><span class="sxs-lookup"><span data-stu-id="a1c20-207">Requirement</span></span>| <span data-ttu-id="a1c20-208">值</span><span class="sxs-lookup"><span data-stu-id="a1c20-208">Value</span></span>|
|---|---|
|[<span data-ttu-id="a1c20-209">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="a1c20-209">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a1c20-210">1.0</span><span class="sxs-lookup"><span data-stu-id="a1c20-210">1.0</span></span>|
|[<span data-ttu-id="a1c20-211">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="a1c20-211">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a1c20-212">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a1c20-212">ReadItem</span></span>|
|[<span data-ttu-id="a1c20-213">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="a1c20-213">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a1c20-214">阅读</span><span class="sxs-lookup"><span data-stu-id="a1c20-214">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a1c20-215">示例</span><span class="sxs-lookup"><span data-stu-id="a1c20-215">Example</span></span>

```JavaScript
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="a1c20-216">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="a1c20-216">dateTimeModified :Date</span></span>

<span data-ttu-id="a1c20-p111">获取项目最近一次修改的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="a1c20-p111">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="a1c20-219">此成员不支持适用于 iOS 的 Outlook 或 Outlook for Android。</span><span class="sxs-lookup"><span data-stu-id="a1c20-219">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="a1c20-220">类型:</span><span class="sxs-lookup"><span data-stu-id="a1c20-220">Type:</span></span>

*   <span data-ttu-id="a1c20-221">日期</span><span class="sxs-lookup"><span data-stu-id="a1c20-221">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="a1c20-222">要求</span><span class="sxs-lookup"><span data-stu-id="a1c20-222">Requirements</span></span>

|<span data-ttu-id="a1c20-223">要求</span><span class="sxs-lookup"><span data-stu-id="a1c20-223">Requirement</span></span>| <span data-ttu-id="a1c20-224">值</span><span class="sxs-lookup"><span data-stu-id="a1c20-224">Value</span></span>|
|---|---|
|[<span data-ttu-id="a1c20-225">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="a1c20-225">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a1c20-226">1.0</span><span class="sxs-lookup"><span data-stu-id="a1c20-226">1.0</span></span>|
|[<span data-ttu-id="a1c20-227">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="a1c20-227">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a1c20-228">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a1c20-228">ReadItem</span></span>|
|[<span data-ttu-id="a1c20-229">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="a1c20-229">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a1c20-230">阅读</span><span class="sxs-lookup"><span data-stu-id="a1c20-230">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a1c20-231">示例</span><span class="sxs-lookup"><span data-stu-id="a1c20-231">Example</span></span>

```JavaScript
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook11officetime"></a><span data-ttu-id="a1c20-232">end :Date|[Time](/javascript/api/outlook_1_1/office.time)</span><span class="sxs-lookup"><span data-stu-id="a1c20-232">end :Date|[Time](/javascript/api/outlook_1_1/office.time)</span></span>

<span data-ttu-id="a1c20-233">获取或设置约会结束的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="a1c20-233">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="a1c20-p112">将 `end` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook11officelocalclienttime) 方法将 end 属性值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="a1c20-p112">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook11officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a1c20-236">阅读模式</span><span class="sxs-lookup"><span data-stu-id="a1c20-236">Read mode</span></span>

<span data-ttu-id="a1c20-237">`end` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="a1c20-237">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="a1c20-238">撰写模式</span><span class="sxs-lookup"><span data-stu-id="a1c20-238">Compose mode</span></span>

<span data-ttu-id="a1c20-239">`end` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="a1c20-239">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="a1c20-240">使用 [`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) 方法设置结束时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="a1c20-240">When you use the [`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="a1c20-241">类型：</span><span class="sxs-lookup"><span data-stu-id="a1c20-241">Type:</span></span>

*   <span data-ttu-id="a1c20-242">Date | [Time](/javascript/api/outlook_1_1/office.time)</span><span class="sxs-lookup"><span data-stu-id="a1c20-242">Date | [Time](/javascript/api/outlook_1_1/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a1c20-243">要求</span><span class="sxs-lookup"><span data-stu-id="a1c20-243">Requirements</span></span>

|<span data-ttu-id="a1c20-244">要求</span><span class="sxs-lookup"><span data-stu-id="a1c20-244">Requirement</span></span>| <span data-ttu-id="a1c20-245">值</span><span class="sxs-lookup"><span data-stu-id="a1c20-245">Value</span></span>|
|---|---|
|[<span data-ttu-id="a1c20-246">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="a1c20-246">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a1c20-247">1.0</span><span class="sxs-lookup"><span data-stu-id="a1c20-247">1.0</span></span>|
|[<span data-ttu-id="a1c20-248">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="a1c20-248">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a1c20-249">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a1c20-249">ReadItem</span></span>|
|[<span data-ttu-id="a1c20-250">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="a1c20-250">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a1c20-251">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="a1c20-251">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a1c20-252">示例</span><span class="sxs-lookup"><span data-stu-id="a1c20-252">Example</span></span>

<span data-ttu-id="a1c20-253">以下示例通过使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) 方法，设置撰写模式下约会的结束时间。</span><span class="sxs-lookup"><span data-stu-id="a1c20-253">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```JavaScript
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

#### <a name="from-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails"></a><span data-ttu-id="a1c20-254">from :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="a1c20-254">from :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span></span>

<span data-ttu-id="a1c20-p113">获取邮件发件人的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="a1c20-p113">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="a1c20-p114">`from` 和 [`sender`](#sender-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails) 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="a1c20-p114">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="a1c20-259">`recipientType`属性`EmailAddressDetails`对象在`from`属性是`undefined`。</span><span class="sxs-lookup"><span data-stu-id="a1c20-259">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="a1c20-260">类型:</span><span class="sxs-lookup"><span data-stu-id="a1c20-260">Type:</span></span>

*   [<span data-ttu-id="a1c20-261">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="a1c20-261">EmailAddressDetails</span></span>](/javascript/api/outlook_1_1/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="a1c20-262">要求</span><span class="sxs-lookup"><span data-stu-id="a1c20-262">Requirements</span></span>

|<span data-ttu-id="a1c20-263">要求</span><span class="sxs-lookup"><span data-stu-id="a1c20-263">Requirement</span></span>| <span data-ttu-id="a1c20-264">值</span><span class="sxs-lookup"><span data-stu-id="a1c20-264">Value</span></span>|
|---|---|
|[<span data-ttu-id="a1c20-265">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="a1c20-265">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a1c20-266">1.0</span><span class="sxs-lookup"><span data-stu-id="a1c20-266">1.0</span></span>|
|[<span data-ttu-id="a1c20-267">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="a1c20-267">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a1c20-268">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a1c20-268">ReadItem</span></span>|
|[<span data-ttu-id="a1c20-269">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="a1c20-269">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a1c20-270">阅读</span><span class="sxs-lookup"><span data-stu-id="a1c20-270">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="a1c20-271">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="a1c20-271">internetMessageId :String</span></span>

<span data-ttu-id="a1c20-p115">获取电子邮件的 Internet 消息标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="a1c20-p115">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="a1c20-274">类型：</span><span class="sxs-lookup"><span data-stu-id="a1c20-274">Type:</span></span>

*   <span data-ttu-id="a1c20-275">String</span><span class="sxs-lookup"><span data-stu-id="a1c20-275">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a1c20-276">要求</span><span class="sxs-lookup"><span data-stu-id="a1c20-276">Requirements</span></span>

|<span data-ttu-id="a1c20-277">要求</span><span class="sxs-lookup"><span data-stu-id="a1c20-277">Requirement</span></span>| <span data-ttu-id="a1c20-278">值</span><span class="sxs-lookup"><span data-stu-id="a1c20-278">Value</span></span>|
|---|---|
|[<span data-ttu-id="a1c20-279">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="a1c20-279">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a1c20-280">1.0</span><span class="sxs-lookup"><span data-stu-id="a1c20-280">1.0</span></span>|
|[<span data-ttu-id="a1c20-281">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="a1c20-281">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a1c20-282">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a1c20-282">ReadItem</span></span>|
|[<span data-ttu-id="a1c20-283">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="a1c20-283">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a1c20-284">阅读</span><span class="sxs-lookup"><span data-stu-id="a1c20-284">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a1c20-285">示例</span><span class="sxs-lookup"><span data-stu-id="a1c20-285">Example</span></span>

```JavaScript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="a1c20-286">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="a1c20-286">itemClass :String</span></span>

<span data-ttu-id="a1c20-p116">获取选定项目的 Exchange Web 服务项目类。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="a1c20-p116">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="a1c20-p117">`itemClass` 属性指定所选项目的邮件类别。以下是邮件或约会项目的默认邮件类别。</span><span class="sxs-lookup"><span data-stu-id="a1c20-p117">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="a1c20-291">类型</span><span class="sxs-lookup"><span data-stu-id="a1c20-291">Type</span></span> | <span data-ttu-id="a1c20-292">说明</span><span class="sxs-lookup"><span data-stu-id="a1c20-292">Description</span></span> | <span data-ttu-id="a1c20-293">项目类</span><span class="sxs-lookup"><span data-stu-id="a1c20-293">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="a1c20-294">约会项目</span><span class="sxs-lookup"><span data-stu-id="a1c20-294">Appointment items</span></span> | <span data-ttu-id="a1c20-295">这些是项目类 `IPM.Appointment` 或 `IPM.Appointment.Occurence` 的日历项目。</span><span class="sxs-lookup"><span data-stu-id="a1c20-295">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| <span data-ttu-id="a1c20-296">邮件项目</span><span class="sxs-lookup"><span data-stu-id="a1c20-296">Message items</span></span> | <span data-ttu-id="a1c20-297">这些项目包括具有默认邮件类别 `IPM.Note` 的电子邮件，以及使用 `IPM.Schedule.Meeting` 作为基础邮件类别的会议请求、响应和取消。</span><span class="sxs-lookup"><span data-stu-id="a1c20-297">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="a1c20-298">你可以创建用于扩展默认邮件类别的自定义邮件类别，例如，自定义约会邮件类别 `IPM.Appointment.Contoso`。</span><span class="sxs-lookup"><span data-stu-id="a1c20-298">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="a1c20-299">类型：</span><span class="sxs-lookup"><span data-stu-id="a1c20-299">Type:</span></span>

*   <span data-ttu-id="a1c20-300">String</span><span class="sxs-lookup"><span data-stu-id="a1c20-300">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a1c20-301">要求</span><span class="sxs-lookup"><span data-stu-id="a1c20-301">Requirements</span></span>

|<span data-ttu-id="a1c20-302">要求</span><span class="sxs-lookup"><span data-stu-id="a1c20-302">Requirement</span></span>| <span data-ttu-id="a1c20-303">值</span><span class="sxs-lookup"><span data-stu-id="a1c20-303">Value</span></span>|
|---|---|
|[<span data-ttu-id="a1c20-304">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="a1c20-304">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a1c20-305">1.0</span><span class="sxs-lookup"><span data-stu-id="a1c20-305">1.0</span></span>|
|[<span data-ttu-id="a1c20-306">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="a1c20-306">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a1c20-307">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a1c20-307">ReadItem</span></span>|
|[<span data-ttu-id="a1c20-308">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="a1c20-308">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a1c20-309">阅读</span><span class="sxs-lookup"><span data-stu-id="a1c20-309">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a1c20-310">示例</span><span class="sxs-lookup"><span data-stu-id="a1c20-310">Example</span></span>

```JavaScript
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="a1c20-311">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="a1c20-311">(nullable) itemId :String</span></span>

<span data-ttu-id="a1c20-p118">获取当前项目的 Exchange Web 服务项目标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="a1c20-p118">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="a1c20-314">`itemId` 属性返回的标识符与 Exchange Web 服务项目标识符相同。</span><span class="sxs-lookup"><span data-stu-id="a1c20-314">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="a1c20-315">`itemId`属性不是与 Outlook 条目 ID 或使用 Outlook REST API 的 ID。</span><span class="sxs-lookup"><span data-stu-id="a1c20-315">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="a1c20-316">之前进行 REST API 调用使用此值，它应转换使用`Office.context.mailbox.convertToRestId`，这是可从要求设置 1.3。</span><span class="sxs-lookup"><span data-stu-id="a1c20-316">Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`, which is available starting in requirement set 1.3.</span></span> <span data-ttu-id="a1c20-317">有关详细信息，请参阅[使用 Outlook REST Api 从 Outlook 外接程序](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id)。</span><span class="sxs-lookup"><span data-stu-id="a1c20-317">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

##### <a name="type"></a><span data-ttu-id="a1c20-318">类型:</span><span class="sxs-lookup"><span data-stu-id="a1c20-318">Type:</span></span>

*   <span data-ttu-id="a1c20-319">String</span><span class="sxs-lookup"><span data-stu-id="a1c20-319">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a1c20-320">要求</span><span class="sxs-lookup"><span data-stu-id="a1c20-320">Requirements</span></span>

|<span data-ttu-id="a1c20-321">要求</span><span class="sxs-lookup"><span data-stu-id="a1c20-321">Requirement</span></span>| <span data-ttu-id="a1c20-322">值</span><span class="sxs-lookup"><span data-stu-id="a1c20-322">Value</span></span>|
|---|---|
|[<span data-ttu-id="a1c20-323">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="a1c20-323">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a1c20-324">1.0</span><span class="sxs-lookup"><span data-stu-id="a1c20-324">1.0</span></span>|
|[<span data-ttu-id="a1c20-325">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="a1c20-325">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a1c20-326">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a1c20-326">ReadItem</span></span>|
|[<span data-ttu-id="a1c20-327">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="a1c20-327">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a1c20-328">阅读</span><span class="sxs-lookup"><span data-stu-id="a1c20-328">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a1c20-329">示例</span><span class="sxs-lookup"><span data-stu-id="a1c20-329">Example</span></span>

<span data-ttu-id="a1c20-p120">以下代码检查项目标识符是否存在。如果 `itemId` 属性返回 `null` 或 `undefined`，则将项目保存到存储，并从异步结果中获取项目标识符。</span><span class="sxs-lookup"><span data-stu-id="a1c20-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```JavaScript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook11officemailboxenumsitemtype"></a><span data-ttu-id="a1c20-332">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_1/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="a1c20-332">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_1/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="a1c20-333">获取实例表示的项的类型。</span><span class="sxs-lookup"><span data-stu-id="a1c20-333">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="a1c20-334">`itemType` 属性返回其中一个 `ItemType` 枚举值，指示 `item` 对象实例是邮件还是约会。</span><span class="sxs-lookup"><span data-stu-id="a1c20-334">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="a1c20-335">类型：</span><span class="sxs-lookup"><span data-stu-id="a1c20-335">Type:</span></span>

*   [<span data-ttu-id="a1c20-336">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="a1c20-336">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_1/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="a1c20-337">要求</span><span class="sxs-lookup"><span data-stu-id="a1c20-337">Requirements</span></span>

|<span data-ttu-id="a1c20-338">要求</span><span class="sxs-lookup"><span data-stu-id="a1c20-338">Requirement</span></span>| <span data-ttu-id="a1c20-339">值</span><span class="sxs-lookup"><span data-stu-id="a1c20-339">Value</span></span>|
|---|---|
|[<span data-ttu-id="a1c20-340">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="a1c20-340">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a1c20-341">1.0</span><span class="sxs-lookup"><span data-stu-id="a1c20-341">1.0</span></span>|
|[<span data-ttu-id="a1c20-342">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="a1c20-342">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a1c20-343">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a1c20-343">ReadItem</span></span>|
|[<span data-ttu-id="a1c20-344">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="a1c20-344">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a1c20-345">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="a1c20-345">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a1c20-346">示例</span><span class="sxs-lookup"><span data-stu-id="a1c20-346">Example</span></span>

```JavaScript
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook11officelocation"></a><span data-ttu-id="a1c20-347">location :String|[Location](/javascript/api/outlook_1_1/office.location)</span><span class="sxs-lookup"><span data-stu-id="a1c20-347">location :String|[Location](/javascript/api/outlook_1_1/office.location)</span></span>

<span data-ttu-id="a1c20-348">获取或设置约会的位置。</span><span class="sxs-lookup"><span data-stu-id="a1c20-348">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a1c20-349">阅读模式</span><span class="sxs-lookup"><span data-stu-id="a1c20-349">Read mode</span></span>

<span data-ttu-id="a1c20-350">`location` 属性返回一个包含约会位置的字符串。</span><span class="sxs-lookup"><span data-stu-id="a1c20-350">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="a1c20-351">撰写模式</span><span class="sxs-lookup"><span data-stu-id="a1c20-351">Compose mode</span></span>

<span data-ttu-id="a1c20-352">`location` 属性返回一个 `Location` 对象，该对象提供用于获取和设置约会位置的方法。</span><span class="sxs-lookup"><span data-stu-id="a1c20-352">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="a1c20-353">类型：</span><span class="sxs-lookup"><span data-stu-id="a1c20-353">Type:</span></span>

*   <span data-ttu-id="a1c20-354">String | [Location](/javascript/api/outlook_1_1/office.location)</span><span class="sxs-lookup"><span data-stu-id="a1c20-354">String | [Location](/javascript/api/outlook_1_1/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a1c20-355">要求</span><span class="sxs-lookup"><span data-stu-id="a1c20-355">Requirements</span></span>

|<span data-ttu-id="a1c20-356">要求</span><span class="sxs-lookup"><span data-stu-id="a1c20-356">Requirement</span></span>| <span data-ttu-id="a1c20-357">值</span><span class="sxs-lookup"><span data-stu-id="a1c20-357">Value</span></span>|
|---|---|
|[<span data-ttu-id="a1c20-358">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="a1c20-358">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a1c20-359">1.0</span><span class="sxs-lookup"><span data-stu-id="a1c20-359">1.0</span></span>|
|[<span data-ttu-id="a1c20-360">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="a1c20-360">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a1c20-361">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a1c20-361">ReadItem</span></span>|
|[<span data-ttu-id="a1c20-362">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="a1c20-362">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a1c20-363">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="a1c20-363">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a1c20-364">示例</span><span class="sxs-lookup"><span data-stu-id="a1c20-364">Example</span></span>

```JavaScript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="a1c20-365">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="a1c20-365">normalizedSubject :String</span></span>

<span data-ttu-id="a1c20-p121">获取删除了所有前缀（包括 `RE:` 和 `FWD:`）的项目主题。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="a1c20-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="a1c20-p122">normalizedSubject 属性获取包含由电子邮件程序添加的任何标准前缀（如 `RE:` 和 `FW:`）的项目主题。若要获取包含完整前缀的项目主题，请使用 [`subject`](#subject-stringsubjectjavascriptapioutlook11officesubject) 属性。</span><span class="sxs-lookup"><span data-stu-id="a1c20-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook11officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="a1c20-370">类型:</span><span class="sxs-lookup"><span data-stu-id="a1c20-370">Type:</span></span>

*   <span data-ttu-id="a1c20-371">String</span><span class="sxs-lookup"><span data-stu-id="a1c20-371">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a1c20-372">要求</span><span class="sxs-lookup"><span data-stu-id="a1c20-372">Requirements</span></span>

|<span data-ttu-id="a1c20-373">要求</span><span class="sxs-lookup"><span data-stu-id="a1c20-373">Requirement</span></span>| <span data-ttu-id="a1c20-374">值</span><span class="sxs-lookup"><span data-stu-id="a1c20-374">Value</span></span>|
|---|---|
|[<span data-ttu-id="a1c20-375">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="a1c20-375">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a1c20-376">1.0</span><span class="sxs-lookup"><span data-stu-id="a1c20-376">1.0</span></span>|
|[<span data-ttu-id="a1c20-377">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="a1c20-377">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a1c20-378">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a1c20-378">ReadItem</span></span>|
|[<span data-ttu-id="a1c20-379">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="a1c20-379">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a1c20-380">阅读</span><span class="sxs-lookup"><span data-stu-id="a1c20-380">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a1c20-381">示例</span><span class="sxs-lookup"><span data-stu-id="a1c20-381">Example</span></span>

```JavaScript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="a1c20-382">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a1c20-382">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="a1c20-383">提供对事件的可选与会者的访问。</span><span class="sxs-lookup"><span data-stu-id="a1c20-383">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="a1c20-384">对象的类型和访问级别取决于当前项的模式。</span><span class="sxs-lookup"><span data-stu-id="a1c20-384">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a1c20-385">阅读模式</span><span class="sxs-lookup"><span data-stu-id="a1c20-385">Read mode</span></span>

<span data-ttu-id="a1c20-386">`optionalAttendees` 属性返回一个数组，其中包含每个可选与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="a1c20-386">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="a1c20-387">撰写模式</span><span class="sxs-lookup"><span data-stu-id="a1c20-387">Compose mode</span></span>

<span data-ttu-id="a1c20-388">`optionalAttendees`属性返回`Recipients`对象，它提供用于获取或更新会议的可选与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="a1c20-388">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="a1c20-389">类型：</span><span class="sxs-lookup"><span data-stu-id="a1c20-389">Type:</span></span>

*   <span data-ttu-id="a1c20-390">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a1c20-390">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a1c20-391">要求</span><span class="sxs-lookup"><span data-stu-id="a1c20-391">Requirements</span></span>

|<span data-ttu-id="a1c20-392">要求</span><span class="sxs-lookup"><span data-stu-id="a1c20-392">Requirement</span></span>| <span data-ttu-id="a1c20-393">值</span><span class="sxs-lookup"><span data-stu-id="a1c20-393">Value</span></span>|
|---|---|
|[<span data-ttu-id="a1c20-394">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="a1c20-394">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a1c20-395">1.0</span><span class="sxs-lookup"><span data-stu-id="a1c20-395">1.0</span></span>|
|[<span data-ttu-id="a1c20-396">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="a1c20-396">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a1c20-397">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a1c20-397">ReadItem</span></span>|
|[<span data-ttu-id="a1c20-398">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="a1c20-398">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a1c20-399">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="a1c20-399">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a1c20-400">示例</span><span class="sxs-lookup"><span data-stu-id="a1c20-400">Example</span></span>

```JavaScript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails"></a><span data-ttu-id="a1c20-401">organizer :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="a1c20-401">organizer :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span></span>

<span data-ttu-id="a1c20-p124">获取指定会议的会议组织者的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="a1c20-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="a1c20-404">类型:</span><span class="sxs-lookup"><span data-stu-id="a1c20-404">Type:</span></span>

*   [<span data-ttu-id="a1c20-405">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="a1c20-405">EmailAddressDetails</span></span>](/javascript/api/outlook_1_1/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="a1c20-406">要求</span><span class="sxs-lookup"><span data-stu-id="a1c20-406">Requirements</span></span>

|<span data-ttu-id="a1c20-407">要求</span><span class="sxs-lookup"><span data-stu-id="a1c20-407">Requirement</span></span>| <span data-ttu-id="a1c20-408">值</span><span class="sxs-lookup"><span data-stu-id="a1c20-408">Value</span></span>|
|---|---|
|[<span data-ttu-id="a1c20-409">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="a1c20-409">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a1c20-410">1.0</span><span class="sxs-lookup"><span data-stu-id="a1c20-410">1.0</span></span>|
|[<span data-ttu-id="a1c20-411">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="a1c20-411">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a1c20-412">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a1c20-412">ReadItem</span></span>|
|[<span data-ttu-id="a1c20-413">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="a1c20-413">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a1c20-414">阅读</span><span class="sxs-lookup"><span data-stu-id="a1c20-414">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a1c20-415">示例</span><span class="sxs-lookup"><span data-stu-id="a1c20-415">Example</span></span>

```JavaScript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="a1c20-416">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a1c20-416">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="a1c20-417">提供对事件的必需与会者的访问。</span><span class="sxs-lookup"><span data-stu-id="a1c20-417">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="a1c20-418">对象的类型和访问级别取决于当前项的模式。</span><span class="sxs-lookup"><span data-stu-id="a1c20-418">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a1c20-419">阅读模式</span><span class="sxs-lookup"><span data-stu-id="a1c20-419">Read mode</span></span>

<span data-ttu-id="a1c20-420">`requiredAttendees` 属性返回一个数组，其中包含每个必需与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="a1c20-420">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="a1c20-421">撰写模式</span><span class="sxs-lookup"><span data-stu-id="a1c20-421">Compose mode</span></span>

<span data-ttu-id="a1c20-422">`requiredAttendees`属性返回`Recipients`对象，它提供用于获取或更新会议的必需的与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="a1c20-422">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="a1c20-423">类型：</span><span class="sxs-lookup"><span data-stu-id="a1c20-423">Type:</span></span>

*   <span data-ttu-id="a1c20-424">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a1c20-424">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a1c20-425">要求</span><span class="sxs-lookup"><span data-stu-id="a1c20-425">Requirements</span></span>

|<span data-ttu-id="a1c20-426">要求</span><span class="sxs-lookup"><span data-stu-id="a1c20-426">Requirement</span></span>| <span data-ttu-id="a1c20-427">值</span><span class="sxs-lookup"><span data-stu-id="a1c20-427">Value</span></span>|
|---|---|
|[<span data-ttu-id="a1c20-428">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="a1c20-428">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a1c20-429">1.0</span><span class="sxs-lookup"><span data-stu-id="a1c20-429">1.0</span></span>|
|[<span data-ttu-id="a1c20-430">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="a1c20-430">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a1c20-431">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a1c20-431">ReadItem</span></span>|
|[<span data-ttu-id="a1c20-432">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="a1c20-432">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a1c20-433">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="a1c20-433">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a1c20-434">示例</span><span class="sxs-lookup"><span data-stu-id="a1c20-434">Example</span></span>

```JavaScript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails"></a><span data-ttu-id="a1c20-435">sender :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="a1c20-435">sender :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span></span>

<span data-ttu-id="a1c20-p126">获取电子邮件发件人的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="a1c20-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="a1c20-p127">[`from`](#from-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails) 和 `sender` 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="a1c20-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="a1c20-440">`recipientType`属性`EmailAddressDetails`对象在`from`属性是`undefined`。</span><span class="sxs-lookup"><span data-stu-id="a1c20-440">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="a1c20-441">类型:</span><span class="sxs-lookup"><span data-stu-id="a1c20-441">Type:</span></span>

*   [<span data-ttu-id="a1c20-442">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="a1c20-442">EmailAddressDetails</span></span>](/javascript/api/outlook_1_1/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="a1c20-443">要求</span><span class="sxs-lookup"><span data-stu-id="a1c20-443">Requirements</span></span>

|<span data-ttu-id="a1c20-444">要求</span><span class="sxs-lookup"><span data-stu-id="a1c20-444">Requirement</span></span>| <span data-ttu-id="a1c20-445">值</span><span class="sxs-lookup"><span data-stu-id="a1c20-445">Value</span></span>|
|---|---|
|[<span data-ttu-id="a1c20-446">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="a1c20-446">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a1c20-447">1.0</span><span class="sxs-lookup"><span data-stu-id="a1c20-447">1.0</span></span>|
|[<span data-ttu-id="a1c20-448">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="a1c20-448">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a1c20-449">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a1c20-449">ReadItem</span></span>|
|[<span data-ttu-id="a1c20-450">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="a1c20-450">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a1c20-451">阅读</span><span class="sxs-lookup"><span data-stu-id="a1c20-451">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a1c20-452">示例</span><span class="sxs-lookup"><span data-stu-id="a1c20-452">Example</span></span>

```JavaScript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook11officetime"></a><span data-ttu-id="a1c20-453">start :Date|[Time](/javascript/api/outlook_1_1/office.time)</span><span class="sxs-lookup"><span data-stu-id="a1c20-453">start :Date|[Time](/javascript/api/outlook_1_1/office.time)</span></span>

<span data-ttu-id="a1c20-454">获取或设置约会开始的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="a1c20-454">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="a1c20-p128">将 `start` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook11officelocalclienttime) 方法将该值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="a1c20-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook11officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a1c20-457">阅读模式</span><span class="sxs-lookup"><span data-stu-id="a1c20-457">Read mode</span></span>

<span data-ttu-id="a1c20-458">`start` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="a1c20-458">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="a1c20-459">撰写模式</span><span class="sxs-lookup"><span data-stu-id="a1c20-459">Compose mode</span></span>

<span data-ttu-id="a1c20-460">`start` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="a1c20-460">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="a1c20-461">使用 [`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) 方法设置开始时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="a1c20-461">When you use the [`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="a1c20-462">类型：</span><span class="sxs-lookup"><span data-stu-id="a1c20-462">Type:</span></span>

*   <span data-ttu-id="a1c20-463">Date | [Time](/javascript/api/outlook_1_1/office.time)</span><span class="sxs-lookup"><span data-stu-id="a1c20-463">Date | [Time](/javascript/api/outlook_1_1/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a1c20-464">要求</span><span class="sxs-lookup"><span data-stu-id="a1c20-464">Requirements</span></span>

|<span data-ttu-id="a1c20-465">要求</span><span class="sxs-lookup"><span data-stu-id="a1c20-465">Requirement</span></span>| <span data-ttu-id="a1c20-466">值</span><span class="sxs-lookup"><span data-stu-id="a1c20-466">Value</span></span>|
|---|---|
|[<span data-ttu-id="a1c20-467">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="a1c20-467">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a1c20-468">1.0</span><span class="sxs-lookup"><span data-stu-id="a1c20-468">1.0</span></span>|
|[<span data-ttu-id="a1c20-469">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="a1c20-469">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a1c20-470">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a1c20-470">ReadItem</span></span>|
|[<span data-ttu-id="a1c20-471">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="a1c20-471">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a1c20-472">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="a1c20-472">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a1c20-473">示例</span><span class="sxs-lookup"><span data-stu-id="a1c20-473">Example</span></span>

<span data-ttu-id="a1c20-474">以下示例通过使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) 方法，设置撰写模式下约会的开始时间。</span><span class="sxs-lookup"><span data-stu-id="a1c20-474">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```JavaScript
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

####  <a name="subject-stringsubjectjavascriptapioutlook11officesubject"></a><span data-ttu-id="a1c20-475">subject :String|[Subject](/javascript/api/outlook_1_1/office.subject)</span><span class="sxs-lookup"><span data-stu-id="a1c20-475">subject :String|[Subject](/javascript/api/outlook_1_1/office.subject)</span></span>

<span data-ttu-id="a1c20-476">获取或设置显示在项目的主题字段中的说明。</span><span class="sxs-lookup"><span data-stu-id="a1c20-476">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="a1c20-477">`subject` 属性获取或设置由电子邮件服务器发送项目时的整个主题。</span><span class="sxs-lookup"><span data-stu-id="a1c20-477">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a1c20-478">阅读模式</span><span class="sxs-lookup"><span data-stu-id="a1c20-478">Read mode</span></span>

<span data-ttu-id="a1c20-p129">`subject` 属性返回一个字符串。使用 [`normalizedSubject`](#normalizedsubject-string) 属性获取不带任何前导前缀（如 `RE:` 和 `FW:`）的主题。</span><span class="sxs-lookup"><span data-stu-id="a1c20-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="a1c20-481">撰写模式</span><span class="sxs-lookup"><span data-stu-id="a1c20-481">Compose mode</span></span>

<span data-ttu-id="a1c20-482">`subject` 属性返回一个 `Subject` 对象，该对象提供用于获取和设置主题的方法。</span><span class="sxs-lookup"><span data-stu-id="a1c20-482">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```JavaScript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="a1c20-483">类型：</span><span class="sxs-lookup"><span data-stu-id="a1c20-483">Type:</span></span>

*   <span data-ttu-id="a1c20-484">String | [Subject](/javascript/api/outlook_1_1/office.subject)</span><span class="sxs-lookup"><span data-stu-id="a1c20-484">String | [Subject](/javascript/api/outlook_1_1/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a1c20-485">要求</span><span class="sxs-lookup"><span data-stu-id="a1c20-485">Requirements</span></span>

|<span data-ttu-id="a1c20-486">要求</span><span class="sxs-lookup"><span data-stu-id="a1c20-486">Requirement</span></span>| <span data-ttu-id="a1c20-487">值</span><span class="sxs-lookup"><span data-stu-id="a1c20-487">Value</span></span>|
|---|---|
|[<span data-ttu-id="a1c20-488">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="a1c20-488">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a1c20-489">1.0</span><span class="sxs-lookup"><span data-stu-id="a1c20-489">1.0</span></span>|
|[<span data-ttu-id="a1c20-490">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="a1c20-490">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a1c20-491">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a1c20-491">ReadItem</span></span>|
|[<span data-ttu-id="a1c20-492">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="a1c20-492">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a1c20-493">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="a1c20-493">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="a1c20-494">到： 数组。 <[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[收件人](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a1c20-494">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="a1c20-495">提供对邮件的**到**行上的收件人访问。</span><span class="sxs-lookup"><span data-stu-id="a1c20-495">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="a1c20-496">对象的类型和访问级别取决于当前项的模式。</span><span class="sxs-lookup"><span data-stu-id="a1c20-496">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a1c20-497">阅读模式</span><span class="sxs-lookup"><span data-stu-id="a1c20-497">Read mode</span></span>

<span data-ttu-id="a1c20-p131">`to` 属性返回包含邮件的**收件人**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="a1c20-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="a1c20-500">撰写模式</span><span class="sxs-lookup"><span data-stu-id="a1c20-500">Compose mode</span></span>

<span data-ttu-id="a1c20-501">`to`属性返回`Recipients`提供方法用于获取或更新邮件的**到**行上的收件人对象。</span><span class="sxs-lookup"><span data-stu-id="a1c20-501">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="a1c20-502">类型：</span><span class="sxs-lookup"><span data-stu-id="a1c20-502">Type:</span></span>

*   <span data-ttu-id="a1c20-503">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a1c20-503">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a1c20-504">要求</span><span class="sxs-lookup"><span data-stu-id="a1c20-504">Requirements</span></span>

|<span data-ttu-id="a1c20-505">要求</span><span class="sxs-lookup"><span data-stu-id="a1c20-505">Requirement</span></span>| <span data-ttu-id="a1c20-506">值</span><span class="sxs-lookup"><span data-stu-id="a1c20-506">Value</span></span>|
|---|---|
|[<span data-ttu-id="a1c20-507">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="a1c20-507">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a1c20-508">1.0</span><span class="sxs-lookup"><span data-stu-id="a1c20-508">1.0</span></span>|
|[<span data-ttu-id="a1c20-509">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="a1c20-509">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a1c20-510">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a1c20-510">ReadItem</span></span>|
|[<span data-ttu-id="a1c20-511">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="a1c20-511">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a1c20-512">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="a1c20-512">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a1c20-513">示例</span><span class="sxs-lookup"><span data-stu-id="a1c20-513">Example</span></span>

```JavaScript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="a1c20-514">方法</span><span class="sxs-lookup"><span data-stu-id="a1c20-514">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="a1c20-515">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="a1c20-515">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="a1c20-516">将文件作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="a1c20-516">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="a1c20-517">`addFileAttachmentAsync` 方法在指定的 URI 上载文件并将其附加到撰写窗体中的项目。</span><span class="sxs-lookup"><span data-stu-id="a1c20-517">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="a1c20-518">你随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="a1c20-518">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a1c20-519">参数：</span><span class="sxs-lookup"><span data-stu-id="a1c20-519">Parameters:</span></span>

|<span data-ttu-id="a1c20-520">名称</span><span class="sxs-lookup"><span data-stu-id="a1c20-520">Name</span></span>| <span data-ttu-id="a1c20-521">类型</span><span class="sxs-lookup"><span data-stu-id="a1c20-521">Type</span></span>| <span data-ttu-id="a1c20-522">属性</span><span class="sxs-lookup"><span data-stu-id="a1c20-522">Attributes</span></span>| <span data-ttu-id="a1c20-523">说明</span><span class="sxs-lookup"><span data-stu-id="a1c20-523">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="a1c20-524">字符串</span><span class="sxs-lookup"><span data-stu-id="a1c20-524">String</span></span>||<span data-ttu-id="a1c20-p132">提供附加到邮件或约会的文件的位置的 URI。最大长度为 2048 个字符。</span><span class="sxs-lookup"><span data-stu-id="a1c20-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="a1c20-527">字符串</span><span class="sxs-lookup"><span data-stu-id="a1c20-527">String</span></span>||<span data-ttu-id="a1c20-p133">在附件上载过程中显示的附件名称。最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="a1c20-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="a1c20-530">对象</span><span class="sxs-lookup"><span data-stu-id="a1c20-530">Object</span></span>| <span data-ttu-id="a1c20-531">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="a1c20-531">&lt;optional&gt;</span></span>|<span data-ttu-id="a1c20-532">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="a1c20-532">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="a1c20-533">对象</span><span class="sxs-lookup"><span data-stu-id="a1c20-533">Object</span></span>| <span data-ttu-id="a1c20-534">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="a1c20-534">&lt;optional&gt;</span></span>|<span data-ttu-id="a1c20-535">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="a1c20-535">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="a1c20-536">函数</span><span class="sxs-lookup"><span data-stu-id="a1c20-536">function</span></span>| <span data-ttu-id="a1c20-537">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="a1c20-537">&lt;optional&gt;</span></span>|<span data-ttu-id="a1c20-538">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="a1c20-538">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="a1c20-539">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="a1c20-539">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="a1c20-540">如果上传附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="a1c20-540">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="a1c20-541">错误</span><span class="sxs-lookup"><span data-stu-id="a1c20-541">Errors</span></span>

| <span data-ttu-id="a1c20-542">错误代码</span><span class="sxs-lookup"><span data-stu-id="a1c20-542">Error code</span></span> | <span data-ttu-id="a1c20-543">说明</span><span class="sxs-lookup"><span data-stu-id="a1c20-543">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="a1c20-544">附件大小超过了允许的大小。</span><span class="sxs-lookup"><span data-stu-id="a1c20-544">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="a1c20-545">该附件的扩展名不是允许的扩展名。</span><span class="sxs-lookup"><span data-stu-id="a1c20-545">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="a1c20-546">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="a1c20-546">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="a1c20-547">要求</span><span class="sxs-lookup"><span data-stu-id="a1c20-547">Requirements</span></span>

|<span data-ttu-id="a1c20-548">要求</span><span class="sxs-lookup"><span data-stu-id="a1c20-548">Requirement</span></span>| <span data-ttu-id="a1c20-549">值</span><span class="sxs-lookup"><span data-stu-id="a1c20-549">Value</span></span>|
|---|---|
|[<span data-ttu-id="a1c20-550">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="a1c20-550">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a1c20-551">1.1</span><span class="sxs-lookup"><span data-stu-id="a1c20-551">1.1</span></span>|
|[<span data-ttu-id="a1c20-552">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="a1c20-552">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a1c20-553">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a1c20-553">ReadWriteItem</span></span>|
|[<span data-ttu-id="a1c20-554">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="a1c20-554">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a1c20-555">撰写</span><span class="sxs-lookup"><span data-stu-id="a1c20-555">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="a1c20-556">示例</span><span class="sxs-lookup"><span data-stu-id="a1c20-556">Example</span></span>

```JavaScript
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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="a1c20-557">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="a1c20-557">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="a1c20-558">将 Exchange 项目（如邮件）作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="a1c20-558">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="a1c20-p134">`addItemAttachmentAsync` 方法将包含指定 Exchange 标识符的项目附加到撰写窗体中的项目。如果指定一个回调方法，此方法使用 `asyncResult` 参数调用，该参数包含一个附件标识符或代码，指示附加项目过程中出现的任何错误。可以使用 `options` 参数将状态信息传递给回调方法（如果需要）。</span><span class="sxs-lookup"><span data-stu-id="a1c20-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="a1c20-562">你随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="a1c20-562">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="a1c20-563">如果您 Office 加载项运行的 Outlook Web App 中`addItemAttachmentAsync`方法可以附加到正在编辑; 的项目之外的项目的项目但是，这不受支持，并且不建议使用。</span><span class="sxs-lookup"><span data-stu-id="a1c20-563">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a1c20-564">参数：</span><span class="sxs-lookup"><span data-stu-id="a1c20-564">Parameters:</span></span>

|<span data-ttu-id="a1c20-565">名称</span><span class="sxs-lookup"><span data-stu-id="a1c20-565">Name</span></span>| <span data-ttu-id="a1c20-566">类型</span><span class="sxs-lookup"><span data-stu-id="a1c20-566">Type</span></span>| <span data-ttu-id="a1c20-567">属性</span><span class="sxs-lookup"><span data-stu-id="a1c20-567">Attributes</span></span>| <span data-ttu-id="a1c20-568">说明</span><span class="sxs-lookup"><span data-stu-id="a1c20-568">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="a1c20-569">字符串</span><span class="sxs-lookup"><span data-stu-id="a1c20-569">String</span></span>||<span data-ttu-id="a1c20-p135">要附加的项目的 Exchange 标识符。最大长度为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="a1c20-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="a1c20-572">字符串</span><span class="sxs-lookup"><span data-stu-id="a1c20-572">String</span></span>||<span data-ttu-id="a1c20-p136">要附加的项目的主题。最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="a1c20-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="a1c20-575">对象</span><span class="sxs-lookup"><span data-stu-id="a1c20-575">Object</span></span>| <span data-ttu-id="a1c20-576">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="a1c20-576">&lt;optional&gt;</span></span>|<span data-ttu-id="a1c20-577">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="a1c20-577">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="a1c20-578">对象</span><span class="sxs-lookup"><span data-stu-id="a1c20-578">Object</span></span>| <span data-ttu-id="a1c20-579">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="a1c20-579">&lt;optional&gt;</span></span>|<span data-ttu-id="a1c20-580">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="a1c20-580">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="a1c20-581">函数</span><span class="sxs-lookup"><span data-stu-id="a1c20-581">function</span></span>| <span data-ttu-id="a1c20-582">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="a1c20-582">&lt;optional&gt;</span></span>|<span data-ttu-id="a1c20-583">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="a1c20-583">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="a1c20-584">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="a1c20-584">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="a1c20-585">如果添加附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="a1c20-585">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="a1c20-586">错误</span><span class="sxs-lookup"><span data-stu-id="a1c20-586">Errors</span></span>

| <span data-ttu-id="a1c20-587">错误代码</span><span class="sxs-lookup"><span data-stu-id="a1c20-587">Error code</span></span> | <span data-ttu-id="a1c20-588">说明</span><span class="sxs-lookup"><span data-stu-id="a1c20-588">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="a1c20-589">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="a1c20-589">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="a1c20-590">要求</span><span class="sxs-lookup"><span data-stu-id="a1c20-590">Requirements</span></span>

|<span data-ttu-id="a1c20-591">要求</span><span class="sxs-lookup"><span data-stu-id="a1c20-591">Requirement</span></span>| <span data-ttu-id="a1c20-592">值</span><span class="sxs-lookup"><span data-stu-id="a1c20-592">Value</span></span>|
|---|---|
|[<span data-ttu-id="a1c20-593">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="a1c20-593">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a1c20-594">1.1</span><span class="sxs-lookup"><span data-stu-id="a1c20-594">1.1</span></span>|
|[<span data-ttu-id="a1c20-595">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="a1c20-595">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a1c20-596">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a1c20-596">ReadWriteItem</span></span>|
|[<span data-ttu-id="a1c20-597">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="a1c20-597">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a1c20-598">撰写</span><span class="sxs-lookup"><span data-stu-id="a1c20-598">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="a1c20-599">示例</span><span class="sxs-lookup"><span data-stu-id="a1c20-599">Example</span></span>

<span data-ttu-id="a1c20-600">以下示例将现有的 Outlook 项目添加为名为 `My Attachment` 的附件。</span><span class="sxs-lookup"><span data-stu-id="a1c20-600">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

```JavaScript
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

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="a1c20-601">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="a1c20-601">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="a1c20-602">显示答复窗体，其中包括所选邮件的发件人和所有收件人或所选约会的组织者和所有与会者。</span><span class="sxs-lookup"><span data-stu-id="a1c20-602">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="a1c20-603">在适用于 iOS 的 Outlook 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="a1c20-603">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="a1c20-604">在 Outlook Web App 中，答复窗体显示为包含 3 列视图的弹出式窗体以及包含 2 列或 1 列视图的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="a1c20-604">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="a1c20-605">如果任意字符串参数超出其限制，`displayReplyAllForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="a1c20-605">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

> [!NOTE]
> <span data-ttu-id="a1c20-606">若要向呼叫中包含附件的能力`displayReplyAllForm`不支持在要求集中 1.1。</span><span class="sxs-lookup"><span data-stu-id="a1c20-606">The ability to include attachments in the call to `displayReplyAllForm` is not supported in requirement set 1.1.</span></span> <span data-ttu-id="a1c20-607">附件支持已添加到要求集 1.2 及以上的 `displayReplyAllForm` 中。</span><span class="sxs-lookup"><span data-stu-id="a1c20-607">Attachment support was added to `displayReplyAllForm` in requirement set 1.2 and above.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a1c20-608">参数：</span><span class="sxs-lookup"><span data-stu-id="a1c20-608">Parameters:</span></span>

|<span data-ttu-id="a1c20-609">名称</span><span class="sxs-lookup"><span data-stu-id="a1c20-609">Name</span></span>| <span data-ttu-id="a1c20-610">类型</span><span class="sxs-lookup"><span data-stu-id="a1c20-610">Type</span></span>| <span data-ttu-id="a1c20-611">说明</span><span class="sxs-lookup"><span data-stu-id="a1c20-611">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="a1c20-612">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="a1c20-612">String &#124; Object</span></span>| |<span data-ttu-id="a1c20-p138">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="a1c20-p138">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="a1c20-615">**或**</span><span class="sxs-lookup"><span data-stu-id="a1c20-615">**OR**</span></span><br/><span data-ttu-id="a1c20-p139">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="a1c20-p139">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="a1c20-618">字符串</span><span class="sxs-lookup"><span data-stu-id="a1c20-618">String</span></span> | <span data-ttu-id="a1c20-619">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="a1c20-619">&lt;optional&gt;</span></span> | <span data-ttu-id="a1c20-p140">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="a1c20-p140">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `callback` | <span data-ttu-id="a1c20-622">函数</span><span class="sxs-lookup"><span data-stu-id="a1c20-622">function</span></span> | <span data-ttu-id="a1c20-623">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="a1c20-623">&lt;optional&gt;</span></span> | <span data-ttu-id="a1c20-624">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="a1c20-624">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="a1c20-625">要求</span><span class="sxs-lookup"><span data-stu-id="a1c20-625">Requirements</span></span>

|<span data-ttu-id="a1c20-626">要求</span><span class="sxs-lookup"><span data-stu-id="a1c20-626">Requirement</span></span>| <span data-ttu-id="a1c20-627">值</span><span class="sxs-lookup"><span data-stu-id="a1c20-627">Value</span></span>|
|---|---|
|[<span data-ttu-id="a1c20-628">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="a1c20-628">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a1c20-629">1.0</span><span class="sxs-lookup"><span data-stu-id="a1c20-629">1.0</span></span>|
|[<span data-ttu-id="a1c20-630">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="a1c20-630">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a1c20-631">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a1c20-631">ReadItem</span></span>|
|[<span data-ttu-id="a1c20-632">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="a1c20-632">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a1c20-633">阅读</span><span class="sxs-lookup"><span data-stu-id="a1c20-633">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="a1c20-634">示例</span><span class="sxs-lookup"><span data-stu-id="a1c20-634">Examples</span></span>

<span data-ttu-id="a1c20-635">以下代码将一个字符串传递到 `displayReplyAllForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="a1c20-635">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="a1c20-636">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="a1c20-636">Reply with an empty body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="a1c20-637">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="a1c20-637">Reply with just a body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="a1c20-638">使用正文和回调答复。</span><span class="sxs-lookup"><span data-stu-id="a1c20-638">Reply with a body and a callback.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="displayreplyformformdata"></a><span data-ttu-id="a1c20-639">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="a1c20-639">displayReplyForm(formData)</span></span>

<span data-ttu-id="a1c20-640">显示答复窗体，其中仅包括所选邮件的发件人或所选约会的组织者。</span><span class="sxs-lookup"><span data-stu-id="a1c20-640">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="a1c20-641">在适用于 iOS 的 Outlook 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="a1c20-641">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="a1c20-642">在 Outlook Web App 中，答复窗体显示为包含 3 列视图的弹出式窗体以及包含 2 列或 1 列视图的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="a1c20-642">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="a1c20-643">如果任意字符串参数超出其限制，`displayReplyForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="a1c20-643">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

> [!NOTE]
> <span data-ttu-id="a1c20-644">若要向呼叫中包含附件的能力`displayReplyForm`不支持在要求集中 1.1。</span><span class="sxs-lookup"><span data-stu-id="a1c20-644">The ability to include attachments in the call to `displayReplyForm` is not supported in requirement set 1.1.</span></span> <span data-ttu-id="a1c20-645">附件支持已添加到要求集 1.2 及以上的 `displayReplyForm` 中。</span><span class="sxs-lookup"><span data-stu-id="a1c20-645">Attachment support was added to `displayReplyForm` in requirement set 1.2 and above.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a1c20-646">参数：</span><span class="sxs-lookup"><span data-stu-id="a1c20-646">Parameters:</span></span>

|<span data-ttu-id="a1c20-647">名称</span><span class="sxs-lookup"><span data-stu-id="a1c20-647">Name</span></span>| <span data-ttu-id="a1c20-648">类型</span><span class="sxs-lookup"><span data-stu-id="a1c20-648">Type</span></span>| <span data-ttu-id="a1c20-649">说明</span><span class="sxs-lookup"><span data-stu-id="a1c20-649">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="a1c20-650">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="a1c20-650">String &#124; Object</span></span>| | <span data-ttu-id="a1c20-p142">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="a1c20-p142">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="a1c20-653">**或**</span><span class="sxs-lookup"><span data-stu-id="a1c20-653">**OR**</span></span><br/><span data-ttu-id="a1c20-p143">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="a1c20-p143">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="a1c20-656">字符串</span><span class="sxs-lookup"><span data-stu-id="a1c20-656">String</span></span> | <span data-ttu-id="a1c20-657">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="a1c20-657">&lt;optional&gt;</span></span> | <span data-ttu-id="a1c20-p144">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="a1c20-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `callback` | <span data-ttu-id="a1c20-660">函数</span><span class="sxs-lookup"><span data-stu-id="a1c20-660">function</span></span> | <span data-ttu-id="a1c20-661">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="a1c20-661">&lt;optional&gt;</span></span> | <span data-ttu-id="a1c20-662">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="a1c20-662">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="a1c20-663">要求</span><span class="sxs-lookup"><span data-stu-id="a1c20-663">Requirements</span></span>

|<span data-ttu-id="a1c20-664">要求</span><span class="sxs-lookup"><span data-stu-id="a1c20-664">Requirement</span></span>| <span data-ttu-id="a1c20-665">值</span><span class="sxs-lookup"><span data-stu-id="a1c20-665">Value</span></span>|
|---|---|
|[<span data-ttu-id="a1c20-666">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="a1c20-666">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a1c20-667">1.0</span><span class="sxs-lookup"><span data-stu-id="a1c20-667">1.0</span></span>|
|[<span data-ttu-id="a1c20-668">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="a1c20-668">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a1c20-669">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a1c20-669">ReadItem</span></span>|
|[<span data-ttu-id="a1c20-670">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="a1c20-670">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a1c20-671">阅读</span><span class="sxs-lookup"><span data-stu-id="a1c20-671">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="a1c20-672">示例</span><span class="sxs-lookup"><span data-stu-id="a1c20-672">Examples</span></span>

<span data-ttu-id="a1c20-673">以下代码将一个字符串传递到 `displayReplyForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="a1c20-673">The following code passes a string to the `displayReplyForm` function.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="a1c20-674">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="a1c20-674">Reply with an empty body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="a1c20-675">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="a1c20-675">Reply with just a body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="a1c20-676">使用正文和回调答复。</span><span class="sxs-lookup"><span data-stu-id="a1c20-676">Reply with a body and a callback.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="getentities--entitiesjavascriptapioutlook11officeentities"></a><span data-ttu-id="a1c20-677">getEntities() → {[Entities](/javascript/api/outlook_1_1/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="a1c20-677">getEntities() → {[Entities](/javascript/api/outlook_1_1/office.entities)}</span></span>

<span data-ttu-id="a1c20-678">获取在选定的项的正文中找到的实体。</span><span class="sxs-lookup"><span data-stu-id="a1c20-678">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="a1c20-679">在适用于 iOS 的 Outlook 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="a1c20-679">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="a1c20-680">要求</span><span class="sxs-lookup"><span data-stu-id="a1c20-680">Requirements</span></span>

|<span data-ttu-id="a1c20-681">要求</span><span class="sxs-lookup"><span data-stu-id="a1c20-681">Requirement</span></span>| <span data-ttu-id="a1c20-682">值</span><span class="sxs-lookup"><span data-stu-id="a1c20-682">Value</span></span>|
|---|---|
|[<span data-ttu-id="a1c20-683">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="a1c20-683">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a1c20-684">1.0</span><span class="sxs-lookup"><span data-stu-id="a1c20-684">1.0</span></span>|
|[<span data-ttu-id="a1c20-685">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="a1c20-685">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a1c20-686">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a1c20-686">ReadItem</span></span>|
|[<span data-ttu-id="a1c20-687">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="a1c20-687">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a1c20-688">阅读</span><span class="sxs-lookup"><span data-stu-id="a1c20-688">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a1c20-689">返回：</span><span class="sxs-lookup"><span data-stu-id="a1c20-689">Returns:</span></span>

<span data-ttu-id="a1c20-690">类型：[Entities](/javascript/api/outlook_1_1/office.entities)</span><span class="sxs-lookup"><span data-stu-id="a1c20-690">Type: [Entities](/javascript/api/outlook_1_1/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="a1c20-691">示例</span><span class="sxs-lookup"><span data-stu-id="a1c20-691">Example</span></span>

<span data-ttu-id="a1c20-692">以下示例访问当前项的主体中的联系人实体。</span><span class="sxs-lookup"><span data-stu-id="a1c20-692">The following example accesses the contacts entities in the current item's body.</span></span>

```JavaScript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook11officecontactmeetingsuggestionjavascriptapioutlook11officemeetingsuggestionphonenumberjavascriptapioutlook11officephonenumbertasksuggestionjavascriptapioutlook11officetasksuggestion"></a><span data-ttu-id="a1c20-693">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="a1c20-693">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}</span></span>

<span data-ttu-id="a1c20-694">获取选定的项的正文中找到的指定的实体类型的所有实体的数组。</span><span class="sxs-lookup"><span data-stu-id="a1c20-694">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="a1c20-695">在适用于 iOS 的 Outlook 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="a1c20-695">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a1c20-696">参数：</span><span class="sxs-lookup"><span data-stu-id="a1c20-696">Parameters:</span></span>

|<span data-ttu-id="a1c20-697">名称</span><span class="sxs-lookup"><span data-stu-id="a1c20-697">Name</span></span>| <span data-ttu-id="a1c20-698">类型</span><span class="sxs-lookup"><span data-stu-id="a1c20-698">Type</span></span>| <span data-ttu-id="a1c20-699">说明</span><span class="sxs-lookup"><span data-stu-id="a1c20-699">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="a1c20-700">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="a1c20-700">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_1/office.MailboxEnums.entitytype)|<span data-ttu-id="a1c20-701">EntityType 枚举值之一。</span><span class="sxs-lookup"><span data-stu-id="a1c20-701">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a1c20-702">要求</span><span class="sxs-lookup"><span data-stu-id="a1c20-702">Requirements</span></span>

|<span data-ttu-id="a1c20-703">要求</span><span class="sxs-lookup"><span data-stu-id="a1c20-703">Requirement</span></span>| <span data-ttu-id="a1c20-704">值</span><span class="sxs-lookup"><span data-stu-id="a1c20-704">Value</span></span>|
|---|---|
|[<span data-ttu-id="a1c20-705">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="a1c20-705">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a1c20-706">1.0</span><span class="sxs-lookup"><span data-stu-id="a1c20-706">1.0</span></span>|
|[<span data-ttu-id="a1c20-707">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="a1c20-707">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a1c20-708">受限</span><span class="sxs-lookup"><span data-stu-id="a1c20-708">Restricted</span></span>|
|[<span data-ttu-id="a1c20-709">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="a1c20-709">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a1c20-710">阅读</span><span class="sxs-lookup"><span data-stu-id="a1c20-710">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a1c20-711">返回：</span><span class="sxs-lookup"><span data-stu-id="a1c20-711">Returns:</span></span>

<span data-ttu-id="a1c20-712">如果在 `entityType` 中传递的值不是 `EntityType` 枚举的有效成员，该方法返回 null。</span><span class="sxs-lookup"><span data-stu-id="a1c20-712">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="a1c20-713">如果没有指定类型的实体的存在于项目的正文中，该方法将返回空数组。</span><span class="sxs-lookup"><span data-stu-id="a1c20-713">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="a1c20-714">否则，返回的数组中对象的类型取决于 `entityType` 参数中请求实体的类型。</span><span class="sxs-lookup"><span data-stu-id="a1c20-714">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="a1c20-715">当使用此方法的最低权限级别**受限**时，某些实体类型需要 **ReadItem** 才能进行访问，如下表中所指定。</span><span class="sxs-lookup"><span data-stu-id="a1c20-715">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="a1c20-716">`entityType` 的值</span><span class="sxs-lookup"><span data-stu-id="a1c20-716">Value of `entityType`</span></span> | <span data-ttu-id="a1c20-717">返回的数组中对象的类型</span><span class="sxs-lookup"><span data-stu-id="a1c20-717">Type of objects in returned array</span></span> | <span data-ttu-id="a1c20-718">所需权限级别</span><span class="sxs-lookup"><span data-stu-id="a1c20-718">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="a1c20-719">String</span><span class="sxs-lookup"><span data-stu-id="a1c20-719">String</span></span> | <span data-ttu-id="a1c20-720">**受限**</span><span class="sxs-lookup"><span data-stu-id="a1c20-720">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="a1c20-721">Contact</span><span class="sxs-lookup"><span data-stu-id="a1c20-721">Contact</span></span> | <span data-ttu-id="a1c20-722">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="a1c20-722">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="a1c20-723">String</span><span class="sxs-lookup"><span data-stu-id="a1c20-723">String</span></span> | <span data-ttu-id="a1c20-724">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="a1c20-724">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="a1c20-725">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="a1c20-725">MeetingSuggestion</span></span> | <span data-ttu-id="a1c20-726">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="a1c20-726">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="a1c20-727">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="a1c20-727">PhoneNumber</span></span> | <span data-ttu-id="a1c20-728">**受限**</span><span class="sxs-lookup"><span data-stu-id="a1c20-728">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="a1c20-729">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="a1c20-729">TaskSuggestion</span></span> | <span data-ttu-id="a1c20-730">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="a1c20-730">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="a1c20-731">String</span><span class="sxs-lookup"><span data-stu-id="a1c20-731">String</span></span> | <span data-ttu-id="a1c20-732">**受限**</span><span class="sxs-lookup"><span data-stu-id="a1c20-732">**Restricted**</span></span> |

<span data-ttu-id="a1c20-733">类型：Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="a1c20-733">Type:  Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))></span></span>


##### <a name="example"></a><span data-ttu-id="a1c20-734">示例</span><span class="sxs-lookup"><span data-stu-id="a1c20-734">Example</span></span>

<span data-ttu-id="a1c20-735">下面的示例演示如何访问表示当前项目的正文中的邮政地址的字符串数组。</span><span class="sxs-lookup"><span data-stu-id="a1c20-735">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

```JavaScript
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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook11officecontactmeetingsuggestionjavascriptapioutlook11officemeetingsuggestionphonenumberjavascriptapioutlook11officephonenumbertasksuggestionjavascriptapioutlook11officetasksuggestion"></a><span data-ttu-id="a1c20-736">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="a1c20-736">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}</span></span>

<span data-ttu-id="a1c20-737">返回传递清单 XML 文件中定义的命名筛选器的所选项目中的已知实体。</span><span class="sxs-lookup"><span data-stu-id="a1c20-737">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="a1c20-738">在适用于 iOS 的 Outlook 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="a1c20-738">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="a1c20-739">`getFilteredEntitiesByName` 方法返回匹配在具有指定 `FilterName` 元素值的清单 XML 文件中的 [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) 规则元素中定义的正则表达式的实体。</span><span class="sxs-lookup"><span data-stu-id="a1c20-739">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a1c20-740">参数：</span><span class="sxs-lookup"><span data-stu-id="a1c20-740">Parameters:</span></span>

|<span data-ttu-id="a1c20-741">名称</span><span class="sxs-lookup"><span data-stu-id="a1c20-741">Name</span></span>| <span data-ttu-id="a1c20-742">类型</span><span class="sxs-lookup"><span data-stu-id="a1c20-742">Type</span></span>| <span data-ttu-id="a1c20-743">说明</span><span class="sxs-lookup"><span data-stu-id="a1c20-743">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="a1c20-744">字符串</span><span class="sxs-lookup"><span data-stu-id="a1c20-744">String</span></span>|<span data-ttu-id="a1c20-745">定义筛选器匹配的 `ItemHasKnownEntity` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="a1c20-745">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a1c20-746">要求</span><span class="sxs-lookup"><span data-stu-id="a1c20-746">Requirements</span></span>

|<span data-ttu-id="a1c20-747">要求</span><span class="sxs-lookup"><span data-stu-id="a1c20-747">Requirement</span></span>| <span data-ttu-id="a1c20-748">值</span><span class="sxs-lookup"><span data-stu-id="a1c20-748">Value</span></span>|
|---|---|
|[<span data-ttu-id="a1c20-749">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="a1c20-749">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a1c20-750">1.0</span><span class="sxs-lookup"><span data-stu-id="a1c20-750">1.0</span></span>|
|[<span data-ttu-id="a1c20-751">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="a1c20-751">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a1c20-752">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a1c20-752">ReadItem</span></span>|
|[<span data-ttu-id="a1c20-753">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="a1c20-753">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a1c20-754">阅读</span><span class="sxs-lookup"><span data-stu-id="a1c20-754">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a1c20-755">返回：</span><span class="sxs-lookup"><span data-stu-id="a1c20-755">Returns:</span></span>

<span data-ttu-id="a1c20-p146">如果具有匹配 `name` 参数的 `FilterName` 元素值的清单中没有任何 `ItemHasKnownEntity` 元素，则该方法返回 `null`。如果 `name` 参数匹配清单中的 `ItemHasKnownEntity` 元素，但在匹配的当前项目中没有实体，则该方法返回一个空数组。</span><span class="sxs-lookup"><span data-stu-id="a1c20-p146">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>


<span data-ttu-id="a1c20-758">类型：Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="a1c20-758">Type: Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))></span></span>


#### <a name="getregexmatches--object"></a><span data-ttu-id="a1c20-759">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="a1c20-759">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="a1c20-760">返回所选项目中匹配在清单 XML 文件中定义的正则表达式的字符串值。</span><span class="sxs-lookup"><span data-stu-id="a1c20-760">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="a1c20-761">在适用于 iOS 的 Outlook 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="a1c20-761">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="a1c20-p147">`getRegExMatches` 方法返回匹配在清单 XML 文件中的每个 `ItemHasRegularExpressionMatch` 或 `ItemHasKnownEntity` 规则元素中定义的正则表达式的字符串。对于 `ItemHasRegularExpressionMatch` 规则，匹配字符串必须发生在该规则指定的项目的属性中。`PropertyName` 简单类型定义支持的属性。</span><span class="sxs-lookup"><span data-stu-id="a1c20-p147">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="a1c20-765">例如，考虑一个外接程序清单具有以下 `Rule` 元素：</span><span class="sxs-lookup"><span data-stu-id="a1c20-765">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```JavaScript
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="a1c20-766">从 `getRegExMatches` 返回的对象应有两个属性：`fruits` 和 `veggies`。</span><span class="sxs-lookup"><span data-stu-id="a1c20-766">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```JavaScript
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

> [!NOTE]
> <span data-ttu-id="a1c20-p148">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。</span><span class="sxs-lookup"><span data-stu-id="a1c20-p148">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="requirements"></a><span data-ttu-id="a1c20-769">要求</span><span class="sxs-lookup"><span data-stu-id="a1c20-769">Requirements</span></span>

|<span data-ttu-id="a1c20-770">要求</span><span class="sxs-lookup"><span data-stu-id="a1c20-770">Requirement</span></span>| <span data-ttu-id="a1c20-771">值</span><span class="sxs-lookup"><span data-stu-id="a1c20-771">Value</span></span>|
|---|---|
|[<span data-ttu-id="a1c20-772">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="a1c20-772">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a1c20-773">1.0</span><span class="sxs-lookup"><span data-stu-id="a1c20-773">1.0</span></span>|
|[<span data-ttu-id="a1c20-774">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="a1c20-774">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a1c20-775">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a1c20-775">ReadItem</span></span>|
|[<span data-ttu-id="a1c20-776">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="a1c20-776">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a1c20-777">阅读</span><span class="sxs-lookup"><span data-stu-id="a1c20-777">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a1c20-778">返回：</span><span class="sxs-lookup"><span data-stu-id="a1c20-778">Returns:</span></span>

<span data-ttu-id="a1c20-p149">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串数组的对象。每个数组的名称等于匹配 `ItemHasRegularExpressionMatch` 规则的 `RegExName` 属性或匹配 `ItemHasKnownEntity` 规则的 `FilterName` 属性的相应值。</span><span class="sxs-lookup"><span data-stu-id="a1c20-p149">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="a1c20-781">

<dt>
类型</dt>


</span><span class="sxs-lookup"><span data-stu-id="a1c20-781">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="a1c20-782">对象</span><span class="sxs-lookup"><span data-stu-id="a1c20-782">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="a1c20-783">示例</span><span class="sxs-lookup"><span data-stu-id="a1c20-783">Example</span></span>

<span data-ttu-id="a1c20-784">以下示例显示了如何访问正则表达式 <rule> 元素 `fruits` 和 `veggies` 的匹配项的数组，这些元素在清单中指定。</rule></span><span class="sxs-lookup"><span data-stu-id="a1c20-784">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```JavaScript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="a1c20-785">getregexmatchesbyname （name） → (nullable) {数组。 < 字符串 >}</span><span class="sxs-lookup"><span data-stu-id="a1c20-785">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="a1c20-786">返回匹配在清单 XML 文件中定义的命名正则表达式的所选项目中的字符串值。</span><span class="sxs-lookup"><span data-stu-id="a1c20-786">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="a1c20-787">在适用于 iOS 的 Outlook 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="a1c20-787">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="a1c20-788">`getRegExMatchesByName` 方法返回匹配在具有指定 `RegExName` 元素值的清单 XML 文件中的 `ItemHasRegularExpressionMatch` 规则元素中定义的正则表达式的字符串。</span><span class="sxs-lookup"><span data-stu-id="a1c20-788">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="a1c20-p150">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。</span><span class="sxs-lookup"><span data-stu-id="a1c20-p150">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a1c20-791">参数：</span><span class="sxs-lookup"><span data-stu-id="a1c20-791">Parameters:</span></span>

|<span data-ttu-id="a1c20-792">名称</span><span class="sxs-lookup"><span data-stu-id="a1c20-792">Name</span></span>| <span data-ttu-id="a1c20-793">类型</span><span class="sxs-lookup"><span data-stu-id="a1c20-793">Type</span></span>| <span data-ttu-id="a1c20-794">说明</span><span class="sxs-lookup"><span data-stu-id="a1c20-794">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="a1c20-795">字符串</span><span class="sxs-lookup"><span data-stu-id="a1c20-795">String</span></span>|<span data-ttu-id="a1c20-796">定义筛选器匹配的 `ItemHasRegularExpressionMatch` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="a1c20-796">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a1c20-797">要求</span><span class="sxs-lookup"><span data-stu-id="a1c20-797">Requirements</span></span>

|<span data-ttu-id="a1c20-798">要求</span><span class="sxs-lookup"><span data-stu-id="a1c20-798">Requirement</span></span>| <span data-ttu-id="a1c20-799">值</span><span class="sxs-lookup"><span data-stu-id="a1c20-799">Value</span></span>|
|---|---|
|[<span data-ttu-id="a1c20-800">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="a1c20-800">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a1c20-801">1.0</span><span class="sxs-lookup"><span data-stu-id="a1c20-801">1.0</span></span>|
|[<span data-ttu-id="a1c20-802">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="a1c20-802">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a1c20-803">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a1c20-803">ReadItem</span></span>|
|[<span data-ttu-id="a1c20-804">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="a1c20-804">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a1c20-805">阅读</span><span class="sxs-lookup"><span data-stu-id="a1c20-805">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a1c20-806">返回：</span><span class="sxs-lookup"><span data-stu-id="a1c20-806">Returns:</span></span>

<span data-ttu-id="a1c20-807">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串的数组。</span><span class="sxs-lookup"><span data-stu-id="a1c20-807">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="a1c20-808">

<dt>类型</dt>

</span><span class="sxs-lookup"><span data-stu-id="a1c20-808">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="a1c20-809">数组。 < 字符串 ></span><span class="sxs-lookup"><span data-stu-id="a1c20-809">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="a1c20-810">示例</span><span class="sxs-lookup"><span data-stu-id="a1c20-810">Example</span></span>

```JavaScript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="a1c20-811">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="a1c20-811">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="a1c20-812">异步加载所选项目上此外接程序的自定义属性。</span><span class="sxs-lookup"><span data-stu-id="a1c20-812">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="a1c20-p151">自定义属性基于每个应用、每个项目存储为键/值对。此方法在回调中返回 `CustomProperties` 对象，该回调提供访问特定于当前项目和当前外接程序的自定义属性的方法。自定义属性未在项目上加密，因此这不应用作安全存储。</span><span class="sxs-lookup"><span data-stu-id="a1c20-p151">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a1c20-816">参数：</span><span class="sxs-lookup"><span data-stu-id="a1c20-816">Parameters:</span></span>

|<span data-ttu-id="a1c20-817">名称</span><span class="sxs-lookup"><span data-stu-id="a1c20-817">Name</span></span>| <span data-ttu-id="a1c20-818">类型</span><span class="sxs-lookup"><span data-stu-id="a1c20-818">Type</span></span>| <span data-ttu-id="a1c20-819">属性</span><span class="sxs-lookup"><span data-stu-id="a1c20-819">Attributes</span></span>| <span data-ttu-id="a1c20-820">说明</span><span class="sxs-lookup"><span data-stu-id="a1c20-820">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="a1c20-821">函数</span><span class="sxs-lookup"><span data-stu-id="a1c20-821">function</span></span>||<span data-ttu-id="a1c20-822">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="a1c20-822">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="a1c20-823">自定义属性作为 `asyncResult.value` 属性中的 [`CustomProperties`](/javascript/api/outlook_1_1/office.customproperties) 对象提供。</span><span class="sxs-lookup"><span data-stu-id="a1c20-823">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_1/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="a1c20-824">此对象可用于获取、 设置，和从项目中删除自定义属性并将更改保存到自定义属性设置回服务器。</span><span class="sxs-lookup"><span data-stu-id="a1c20-824">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="a1c20-825">对象</span><span class="sxs-lookup"><span data-stu-id="a1c20-825">Object</span></span>| <span data-ttu-id="a1c20-826">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="a1c20-826">&lt;optional&gt;</span></span>|<span data-ttu-id="a1c20-827">开发人员可以提供他们想要访问的回调函数中的任何对象。</span><span class="sxs-lookup"><span data-stu-id="a1c20-827">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="a1c20-828">此对象可以访问`asyncResult.asyncContext`的回调函数中的属性。</span><span class="sxs-lookup"><span data-stu-id="a1c20-828">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a1c20-829">要求</span><span class="sxs-lookup"><span data-stu-id="a1c20-829">Requirements</span></span>

|<span data-ttu-id="a1c20-830">要求</span><span class="sxs-lookup"><span data-stu-id="a1c20-830">Requirement</span></span>| <span data-ttu-id="a1c20-831">值</span><span class="sxs-lookup"><span data-stu-id="a1c20-831">Value</span></span>|
|---|---|
|[<span data-ttu-id="a1c20-832">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="a1c20-832">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a1c20-833">1.0</span><span class="sxs-lookup"><span data-stu-id="a1c20-833">1.0</span></span>|
|[<span data-ttu-id="a1c20-834">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="a1c20-834">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a1c20-835">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a1c20-835">ReadItem</span></span>|
|[<span data-ttu-id="a1c20-836">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="a1c20-836">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a1c20-837">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="a1c20-837">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a1c20-838">示例</span><span class="sxs-lookup"><span data-stu-id="a1c20-838">Example</span></span>

<span data-ttu-id="a1c20-p154">以下代码示例显示了如何使用 `loadCustomPropertiesAsync` 方法异步加载特定于当前项目的自定义属性。该示例还显示了如何使用 `CustomProperties.saveAsync` 方法将这些属性重新保存到服务器。加载自定义属性后，该代码示例将使用 `CustomProperties.get` 方法读取自定义属性 `myProp`，使用 `CustomProperties.set` 方法写入自定义属性 `otherProp`，最后调用 `saveAsync` 方法保存这些自定义属性。</span><span class="sxs-lookup"><span data-stu-id="a1c20-p154">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

```JavaScript
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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="a1c20-842">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="a1c20-842">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="a1c20-843">将附件从邮件或约会中删除。</span><span class="sxs-lookup"><span data-stu-id="a1c20-843">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="a1c20-p155">`removeAttachmentAsync` 方法删除项目中带指定标识符的附件。最佳做法是，仅当同一个邮件应用程序在同一会话中添加了一个附件时，你才应使用该附件标识符来删除该附件。在 Outlook Web App 和适用于设备的 OWA 中，附件标识符只在同一个会话中才有效。当用户关闭应用程序，或者如果用户开始在内嵌窗体中撰写，并在随后弹出的内嵌窗体中继续在单独的窗口撰写时，会话即结束。</span><span class="sxs-lookup"><span data-stu-id="a1c20-p155">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a1c20-848">参数：</span><span class="sxs-lookup"><span data-stu-id="a1c20-848">Parameters:</span></span>

|<span data-ttu-id="a1c20-849">名称</span><span class="sxs-lookup"><span data-stu-id="a1c20-849">Name</span></span>| <span data-ttu-id="a1c20-850">类型</span><span class="sxs-lookup"><span data-stu-id="a1c20-850">Type</span></span>| <span data-ttu-id="a1c20-851">属性</span><span class="sxs-lookup"><span data-stu-id="a1c20-851">Attributes</span></span>| <span data-ttu-id="a1c20-852">说明</span><span class="sxs-lookup"><span data-stu-id="a1c20-852">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="a1c20-853">字符串</span><span class="sxs-lookup"><span data-stu-id="a1c20-853">String</span></span>||<span data-ttu-id="a1c20-p156">要删除的附件的标识符。字符串的最大长度为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="a1c20-p156">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`| <span data-ttu-id="a1c20-856">对象</span><span class="sxs-lookup"><span data-stu-id="a1c20-856">Object</span></span>| <span data-ttu-id="a1c20-857">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="a1c20-857">&lt;optional&gt;</span></span>|<span data-ttu-id="a1c20-858">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="a1c20-858">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="a1c20-859">对象</span><span class="sxs-lookup"><span data-stu-id="a1c20-859">Object</span></span>| <span data-ttu-id="a1c20-860">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="a1c20-860">&lt;optional&gt;</span></span>|<span data-ttu-id="a1c20-861">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="a1c20-861">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="a1c20-862">函数</span><span class="sxs-lookup"><span data-stu-id="a1c20-862">function</span></span>| <span data-ttu-id="a1c20-863">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="a1c20-863">&lt;optional&gt;</span></span>|<span data-ttu-id="a1c20-864">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="a1c20-864">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="a1c20-865">如果删除附件失败，`asyncResult.error` 属性将包含一个说明失败原因的错误代码。</span><span class="sxs-lookup"><span data-stu-id="a1c20-865">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="a1c20-866">错误</span><span class="sxs-lookup"><span data-stu-id="a1c20-866">Errors</span></span>

| <span data-ttu-id="a1c20-867">错误代码</span><span class="sxs-lookup"><span data-stu-id="a1c20-867">Error code</span></span> | <span data-ttu-id="a1c20-868">说明</span><span class="sxs-lookup"><span data-stu-id="a1c20-868">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="a1c20-869">附件标识符不存在。</span><span class="sxs-lookup"><span data-stu-id="a1c20-869">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="a1c20-870">要求</span><span class="sxs-lookup"><span data-stu-id="a1c20-870">Requirements</span></span>

|<span data-ttu-id="a1c20-871">要求</span><span class="sxs-lookup"><span data-stu-id="a1c20-871">Requirement</span></span>| <span data-ttu-id="a1c20-872">值</span><span class="sxs-lookup"><span data-stu-id="a1c20-872">Value</span></span>|
|---|---|
|[<span data-ttu-id="a1c20-873">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="a1c20-873">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a1c20-874">1.1</span><span class="sxs-lookup"><span data-stu-id="a1c20-874">1.1</span></span>|
|[<span data-ttu-id="a1c20-875">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="a1c20-875">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a1c20-876">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a1c20-876">ReadWriteItem</span></span>|
|[<span data-ttu-id="a1c20-877">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="a1c20-877">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a1c20-878">撰写</span><span class="sxs-lookup"><span data-stu-id="a1c20-878">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="a1c20-879">示例</span><span class="sxs-lookup"><span data-stu-id="a1c20-879">Example</span></span>

<span data-ttu-id="a1c20-880">以下代码删除包含标识符 '0' 的附件。</span><span class="sxs-lookup"><span data-stu-id="a1c20-880">The following code removes an attachment with an identifier of '0'.</span></span>

```JavaScript
Office.context.mailbox.item.removeAttachmentAsync(
  '0',
  { asyncContext : null },
  function (asyncResult)
  {
    console.log(asyncResult.status);
  }
);
```