 

# <a name="office"></a><span data-ttu-id="3de8d-101">Office</span><span class="sxs-lookup"><span data-stu-id="3de8d-101">Office</span></span>

<span data-ttu-id="3de8d-p101">该 Office 命名空间提供所有 Office 应用中的外接程序所使用的共享接口。此列表仅记录 Outlook 外接程序所使用的接口。有关 Office 命名空间的完整列表，请参阅[共享 API](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="3de8d-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="3de8d-104">Requirements</span><span class="sxs-lookup"><span data-stu-id="3de8d-104">Requirements</span></span>

|<span data-ttu-id="3de8d-105">要求</span><span class="sxs-lookup"><span data-stu-id="3de8d-105">Requirement</span></span>| <span data-ttu-id="3de8d-106">值</span><span class="sxs-lookup"><span data-stu-id="3de8d-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="3de8d-107">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3de8d-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3de8d-108">1.0</span><span class="sxs-lookup"><span data-stu-id="3de8d-108">1.0</span></span>|
|[<span data-ttu-id="3de8d-109">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3de8d-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3de8d-110">Compose 或 Read</span><span class="sxs-lookup"><span data-stu-id="3de8d-110">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="3de8d-111">成员和方法</span><span class="sxs-lookup"><span data-stu-id="3de8d-111">Members and methods</span></span>

| <span data-ttu-id="3de8d-112">成员</span><span class="sxs-lookup"><span data-stu-id="3de8d-112">Member</span></span> | <span data-ttu-id="3de8d-113">类型</span><span class="sxs-lookup"><span data-stu-id="3de8d-113">Type</span></span> |
|--------|------|
| [<span data-ttu-id="3de8d-114">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="3de8d-114">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="3de8d-115">成员</span><span class="sxs-lookup"><span data-stu-id="3de8d-115">Member</span></span> |
| [<span data-ttu-id="3de8d-116">CoercionType</span><span class="sxs-lookup"><span data-stu-id="3de8d-116">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="3de8d-117">成员</span><span class="sxs-lookup"><span data-stu-id="3de8d-117">Member</span></span> |
| [<span data-ttu-id="3de8d-118">EventType</span><span class="sxs-lookup"><span data-stu-id="3de8d-118">EventType</span></span>](#eventtype-string) | <span data-ttu-id="3de8d-119">成员</span><span class="sxs-lookup"><span data-stu-id="3de8d-119">Member</span></span> |
| [<span data-ttu-id="3de8d-120">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="3de8d-120">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="3de8d-121">成员</span><span class="sxs-lookup"><span data-stu-id="3de8d-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="3de8d-122">命名空间</span><span class="sxs-lookup"><span data-stu-id="3de8d-122">Namespaces</span></span>

<span data-ttu-id="3de8d-123">[上下文](office.context.md)： 用于 Outlook 外接程序 API 提供共享的 Office 加载项 API 上下文命名空间中的接口。</span><span class="sxs-lookup"><span data-stu-id="3de8d-123">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="3de8d-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype)：包括 ItemType、EntityType、AttachmentType、RecipientType、ResponseType 和 ItemNotificationMessageType 枚举。</span><span class="sxs-lookup"><span data-stu-id="3de8d-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="3de8d-125">成员</span><span class="sxs-lookup"><span data-stu-id="3de8d-125">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="3de8d-126">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="3de8d-126">AsyncResultStatus :String</span></span>

<span data-ttu-id="3de8d-127">指定异步调用的结果。</span><span class="sxs-lookup"><span data-stu-id="3de8d-127">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="3de8d-128">类型：</span><span class="sxs-lookup"><span data-stu-id="3de8d-128">Type:</span></span>

*   <span data-ttu-id="3de8d-129">字符串</span><span class="sxs-lookup"><span data-stu-id="3de8d-129">String</span></span>

##### <a name="properties"></a><span data-ttu-id="3de8d-130">属性：</span><span class="sxs-lookup"><span data-stu-id="3de8d-130">Properties:</span></span>

|<span data-ttu-id="3de8d-131">名称</span><span class="sxs-lookup"><span data-stu-id="3de8d-131">Name</span></span>| <span data-ttu-id="3de8d-132">类型</span><span class="sxs-lookup"><span data-stu-id="3de8d-132">Type</span></span>| <span data-ttu-id="3de8d-133">说明</span><span class="sxs-lookup"><span data-stu-id="3de8d-133">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="3de8d-134">String</span><span class="sxs-lookup"><span data-stu-id="3de8d-134">String</span></span>|<span data-ttu-id="3de8d-135">调用成功。</span><span class="sxs-lookup"><span data-stu-id="3de8d-135">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="3de8d-136">字符串</span><span class="sxs-lookup"><span data-stu-id="3de8d-136">String</span></span>|<span data-ttu-id="3de8d-137">调用失败。</span><span class="sxs-lookup"><span data-stu-id="3de8d-137">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3de8d-138">要求</span><span class="sxs-lookup"><span data-stu-id="3de8d-138">Requirements</span></span>

|<span data-ttu-id="3de8d-139">要求</span><span class="sxs-lookup"><span data-stu-id="3de8d-139">Requirement</span></span>| <span data-ttu-id="3de8d-140">值</span><span class="sxs-lookup"><span data-stu-id="3de8d-140">Value</span></span>|
|---|---|
|[<span data-ttu-id="3de8d-141">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3de8d-141">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3de8d-142">1.0</span><span class="sxs-lookup"><span data-stu-id="3de8d-142">1.0</span></span>|
|[<span data-ttu-id="3de8d-143">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3de8d-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3de8d-144">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3de8d-144">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="3de8d-145">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="3de8d-145">CoercionType :String</span></span>

<span data-ttu-id="3de8d-146">指定如何强制由调用方法返回或设置的数据。</span><span class="sxs-lookup"><span data-stu-id="3de8d-146">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="3de8d-147">类型：</span><span class="sxs-lookup"><span data-stu-id="3de8d-147">Type:</span></span>

*   <span data-ttu-id="3de8d-148">字符串</span><span class="sxs-lookup"><span data-stu-id="3de8d-148">String</span></span>

##### <a name="properties"></a><span data-ttu-id="3de8d-149">属性：</span><span class="sxs-lookup"><span data-stu-id="3de8d-149">Properties:</span></span>

|<span data-ttu-id="3de8d-150">名称</span><span class="sxs-lookup"><span data-stu-id="3de8d-150">Name</span></span>| <span data-ttu-id="3de8d-151">类型</span><span class="sxs-lookup"><span data-stu-id="3de8d-151">Type</span></span>| <span data-ttu-id="3de8d-152">说明</span><span class="sxs-lookup"><span data-stu-id="3de8d-152">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="3de8d-153">String</span><span class="sxs-lookup"><span data-stu-id="3de8d-153">String</span></span>|<span data-ttu-id="3de8d-154">请求以 HTML 格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="3de8d-154">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="3de8d-155">字符串</span><span class="sxs-lookup"><span data-stu-id="3de8d-155">String</span></span>|<span data-ttu-id="3de8d-156">请求以文本格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="3de8d-156">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3de8d-157">要求</span><span class="sxs-lookup"><span data-stu-id="3de8d-157">Requirements</span></span>

|<span data-ttu-id="3de8d-158">要求</span><span class="sxs-lookup"><span data-stu-id="3de8d-158">Requirement</span></span>| <span data-ttu-id="3de8d-159">值</span><span class="sxs-lookup"><span data-stu-id="3de8d-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="3de8d-160">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3de8d-160">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3de8d-161">1.0</span><span class="sxs-lookup"><span data-stu-id="3de8d-161">1.0</span></span>|
|[<span data-ttu-id="3de8d-162">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3de8d-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3de8d-163">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3de8d-163">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="3de8d-164">EventType :String</span><span class="sxs-lookup"><span data-stu-id="3de8d-164">EventType :String</span></span>

<span data-ttu-id="3de8d-165">指定与事件处理程序相关联的事件。</span><span class="sxs-lookup"><span data-stu-id="3de8d-165">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="3de8d-166">类型：</span><span class="sxs-lookup"><span data-stu-id="3de8d-166">Type:</span></span>

*   <span data-ttu-id="3de8d-167">字符串</span><span class="sxs-lookup"><span data-stu-id="3de8d-167">String</span></span>

##### <a name="properties"></a><span data-ttu-id="3de8d-168">属性：</span><span class="sxs-lookup"><span data-stu-id="3de8d-168">Properties:</span></span>

| <span data-ttu-id="3de8d-169">名称</span><span class="sxs-lookup"><span data-stu-id="3de8d-169">Name</span></span> | <span data-ttu-id="3de8d-170">类型</span><span class="sxs-lookup"><span data-stu-id="3de8d-170">Type</span></span> | <span data-ttu-id="3de8d-171">说明</span><span class="sxs-lookup"><span data-stu-id="3de8d-171">Description</span></span> | <span data-ttu-id="3de8d-172">最低要求集</span><span class="sxs-lookup"><span data-stu-id="3de8d-172">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="3de8d-173">String</span><span class="sxs-lookup"><span data-stu-id="3de8d-173">String</span></span> | <span data-ttu-id="3de8d-174">已更改的日期或时间所选的约会或系列。</span><span class="sxs-lookup"><span data-stu-id="3de8d-174">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="3de8d-175">1.7</span><span class="sxs-lookup"><span data-stu-id="3de8d-175">1.7</span></span> |
|`ItemChanged`| <span data-ttu-id="3de8d-176">字符串</span><span class="sxs-lookup"><span data-stu-id="3de8d-176">String</span></span> | <span data-ttu-id="3de8d-177">选定的项已更改。</span><span class="sxs-lookup"><span data-stu-id="3de8d-177">The selected item has changed.</span></span> | <span data-ttu-id="3de8d-178">1.5</span><span class="sxs-lookup"><span data-stu-id="3de8d-178">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="3de8d-179">字符串</span><span class="sxs-lookup"><span data-stu-id="3de8d-179">String</span></span> | <span data-ttu-id="3de8d-180">选定的项已更改。</span><span class="sxs-lookup"><span data-stu-id="3de8d-180">The selected item has changed.</span></span> | <span data-ttu-id="3de8d-181">Preview</span><span class="sxs-lookup"><span data-stu-id="3de8d-181">Preview</span></span> |
|`RecipientsChanged`| <span data-ttu-id="3de8d-182">String</span><span class="sxs-lookup"><span data-stu-id="3de8d-182">String</span></span> | <span data-ttu-id="3de8d-183">已更改所选的项目或约会位置的收件人列表。</span><span class="sxs-lookup"><span data-stu-id="3de8d-183">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="3de8d-184">1.7</span><span class="sxs-lookup"><span data-stu-id="3de8d-184">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="3de8d-185">String</span><span class="sxs-lookup"><span data-stu-id="3de8d-185">String</span></span> | <span data-ttu-id="3de8d-186">所选系列的定期模式已更改。</span><span class="sxs-lookup"><span data-stu-id="3de8d-186">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="3de8d-187">1.7</span><span class="sxs-lookup"><span data-stu-id="3de8d-187">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="3de8d-188">要求</span><span class="sxs-lookup"><span data-stu-id="3de8d-188">Requirements</span></span>

|<span data-ttu-id="3de8d-189">要求</span><span class="sxs-lookup"><span data-stu-id="3de8d-189">Requirement</span></span>| <span data-ttu-id="3de8d-190">值</span><span class="sxs-lookup"><span data-stu-id="3de8d-190">Value</span></span>|
|---|---|
|[<span data-ttu-id="3de8d-191">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3de8d-191">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3de8d-192">1.5</span><span class="sxs-lookup"><span data-stu-id="3de8d-192">1.5</span></span> |
|[<span data-ttu-id="3de8d-193">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3de8d-193">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3de8d-194">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3de8d-194">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="3de8d-195">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="3de8d-195">SourceProperty :String</span></span>

<span data-ttu-id="3de8d-196">指定由调用方法返回的数据源。</span><span class="sxs-lookup"><span data-stu-id="3de8d-196">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="3de8d-197">类型：</span><span class="sxs-lookup"><span data-stu-id="3de8d-197">Type:</span></span>

*   <span data-ttu-id="3de8d-198">字符串</span><span class="sxs-lookup"><span data-stu-id="3de8d-198">String</span></span>

##### <a name="properties"></a><span data-ttu-id="3de8d-199">属性：</span><span class="sxs-lookup"><span data-stu-id="3de8d-199">Properties:</span></span>

|<span data-ttu-id="3de8d-200">名称</span><span class="sxs-lookup"><span data-stu-id="3de8d-200">Name</span></span>| <span data-ttu-id="3de8d-201">类型</span><span class="sxs-lookup"><span data-stu-id="3de8d-201">Type</span></span>| <span data-ttu-id="3de8d-202">说明</span><span class="sxs-lookup"><span data-stu-id="3de8d-202">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="3de8d-203">字符串</span><span class="sxs-lookup"><span data-stu-id="3de8d-203">String</span></span>|<span data-ttu-id="3de8d-204">数据源来自邮件的正文。</span><span class="sxs-lookup"><span data-stu-id="3de8d-204">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="3de8d-205">String</span><span class="sxs-lookup"><span data-stu-id="3de8d-205">String</span></span>|<span data-ttu-id="3de8d-206">数据源来自邮件的主题。</span><span class="sxs-lookup"><span data-stu-id="3de8d-206">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3de8d-207">要求</span><span class="sxs-lookup"><span data-stu-id="3de8d-207">Requirements</span></span>

|<span data-ttu-id="3de8d-208">要求</span><span class="sxs-lookup"><span data-stu-id="3de8d-208">Requirement</span></span>| <span data-ttu-id="3de8d-209">值</span><span class="sxs-lookup"><span data-stu-id="3de8d-209">Value</span></span>|
|---|---|
|[<span data-ttu-id="3de8d-210">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3de8d-210">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3de8d-211">1.0</span><span class="sxs-lookup"><span data-stu-id="3de8d-211">1.0</span></span>|
|[<span data-ttu-id="3de8d-212">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3de8d-212">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3de8d-213">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3de8d-213">Compose or read</span></span>|