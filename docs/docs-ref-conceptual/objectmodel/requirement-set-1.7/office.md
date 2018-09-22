 

# <a name="office"></a><span data-ttu-id="6bc25-101">Office</span><span class="sxs-lookup"><span data-stu-id="6bc25-101">Office</span></span>

<span data-ttu-id="6bc25-p101">该 Office 命名空间提供所有 Office 应用中的外接程序所使用的共享接口。此列表仅记录 Outlook 外接程序所使用的接口。有关 Office 命名空间的完整列表，请参阅[共享 API](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="6bc25-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="6bc25-104">Requirements</span><span class="sxs-lookup"><span data-stu-id="6bc25-104">Requirements</span></span>

|<span data-ttu-id="6bc25-105">要求</span><span class="sxs-lookup"><span data-stu-id="6bc25-105">Requirement</span></span>| <span data-ttu-id="6bc25-106">值</span><span class="sxs-lookup"><span data-stu-id="6bc25-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="6bc25-107">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="6bc25-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6bc25-108">1.0</span><span class="sxs-lookup"><span data-stu-id="6bc25-108">1.0</span></span>|
|[<span data-ttu-id="6bc25-109">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="6bc25-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6bc25-110">Compose 或 Read</span><span class="sxs-lookup"><span data-stu-id="6bc25-110">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="6bc25-111">成员和方法</span><span class="sxs-lookup"><span data-stu-id="6bc25-111">Members and methods</span></span>

| <span data-ttu-id="6bc25-112">成员</span><span class="sxs-lookup"><span data-stu-id="6bc25-112">Member</span></span> | <span data-ttu-id="6bc25-113">类型</span><span class="sxs-lookup"><span data-stu-id="6bc25-113">Type</span></span> |
|--------|------|
| [<span data-ttu-id="6bc25-114">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="6bc25-114">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="6bc25-115">成员</span><span class="sxs-lookup"><span data-stu-id="6bc25-115">Member</span></span> |
| [<span data-ttu-id="6bc25-116">CoercionType</span><span class="sxs-lookup"><span data-stu-id="6bc25-116">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="6bc25-117">成员</span><span class="sxs-lookup"><span data-stu-id="6bc25-117">Member</span></span> |
| [<span data-ttu-id="6bc25-118">EventType</span><span class="sxs-lookup"><span data-stu-id="6bc25-118">EventType</span></span>](#eventtype-string) | <span data-ttu-id="6bc25-119">成员</span><span class="sxs-lookup"><span data-stu-id="6bc25-119">Member</span></span> |
| [<span data-ttu-id="6bc25-120">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="6bc25-120">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="6bc25-121">成员</span><span class="sxs-lookup"><span data-stu-id="6bc25-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="6bc25-122">命名空间</span><span class="sxs-lookup"><span data-stu-id="6bc25-122">Namespaces</span></span>

<span data-ttu-id="6bc25-123">[上下文](office.context.md)： 用于 Outlook 外接程序 API 提供共享的 Office 加载项 API 上下文命名空间中的接口。</span><span class="sxs-lookup"><span data-stu-id="6bc25-123">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="6bc25-124">[MailboxEnums](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype)：包括 ItemType、EntityType、AttachmentType、RecipientType、ResponseType 和 ItemNotificationMessageType 枚举。</span><span class="sxs-lookup"><span data-stu-id="6bc25-124">[MailboxEnums](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="6bc25-125">成员</span><span class="sxs-lookup"><span data-stu-id="6bc25-125">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="6bc25-126">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="6bc25-126">AsyncResultStatus :String</span></span>

<span data-ttu-id="6bc25-127">指定异步调用的结果。</span><span class="sxs-lookup"><span data-stu-id="6bc25-127">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="6bc25-128">类型：</span><span class="sxs-lookup"><span data-stu-id="6bc25-128">Type:</span></span>

*   <span data-ttu-id="6bc25-129">字符串</span><span class="sxs-lookup"><span data-stu-id="6bc25-129">String</span></span>

##### <a name="properties"></a><span data-ttu-id="6bc25-130">属性：</span><span class="sxs-lookup"><span data-stu-id="6bc25-130">Properties:</span></span>

|<span data-ttu-id="6bc25-131">名称</span><span class="sxs-lookup"><span data-stu-id="6bc25-131">Name</span></span>| <span data-ttu-id="6bc25-132">类型</span><span class="sxs-lookup"><span data-stu-id="6bc25-132">Type</span></span>| <span data-ttu-id="6bc25-133">说明</span><span class="sxs-lookup"><span data-stu-id="6bc25-133">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="6bc25-134">String</span><span class="sxs-lookup"><span data-stu-id="6bc25-134">String</span></span>|<span data-ttu-id="6bc25-135">调用成功。</span><span class="sxs-lookup"><span data-stu-id="6bc25-135">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="6bc25-136">字符串</span><span class="sxs-lookup"><span data-stu-id="6bc25-136">String</span></span>|<span data-ttu-id="6bc25-137">调用失败。</span><span class="sxs-lookup"><span data-stu-id="6bc25-137">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6bc25-138">要求</span><span class="sxs-lookup"><span data-stu-id="6bc25-138">Requirements</span></span>

|<span data-ttu-id="6bc25-139">要求</span><span class="sxs-lookup"><span data-stu-id="6bc25-139">Requirement</span></span>| <span data-ttu-id="6bc25-140">值</span><span class="sxs-lookup"><span data-stu-id="6bc25-140">Value</span></span>|
|---|---|
|[<span data-ttu-id="6bc25-141">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="6bc25-141">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6bc25-142">1.0</span><span class="sxs-lookup"><span data-stu-id="6bc25-142">1.0</span></span>|
|[<span data-ttu-id="6bc25-143">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="6bc25-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6bc25-144">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="6bc25-144">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="6bc25-145">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="6bc25-145">CoercionType :String</span></span>

<span data-ttu-id="6bc25-146">指定如何强制由调用方法返回或设置的数据。</span><span class="sxs-lookup"><span data-stu-id="6bc25-146">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="6bc25-147">类型：</span><span class="sxs-lookup"><span data-stu-id="6bc25-147">Type:</span></span>

*   <span data-ttu-id="6bc25-148">字符串</span><span class="sxs-lookup"><span data-stu-id="6bc25-148">String</span></span>

##### <a name="properties"></a><span data-ttu-id="6bc25-149">属性：</span><span class="sxs-lookup"><span data-stu-id="6bc25-149">Properties:</span></span>

|<span data-ttu-id="6bc25-150">名称</span><span class="sxs-lookup"><span data-stu-id="6bc25-150">Name</span></span>| <span data-ttu-id="6bc25-151">类型</span><span class="sxs-lookup"><span data-stu-id="6bc25-151">Type</span></span>| <span data-ttu-id="6bc25-152">说明</span><span class="sxs-lookup"><span data-stu-id="6bc25-152">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="6bc25-153">String</span><span class="sxs-lookup"><span data-stu-id="6bc25-153">String</span></span>|<span data-ttu-id="6bc25-154">请求以 HTML 格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="6bc25-154">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="6bc25-155">字符串</span><span class="sxs-lookup"><span data-stu-id="6bc25-155">String</span></span>|<span data-ttu-id="6bc25-156">请求以文本格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="6bc25-156">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6bc25-157">要求</span><span class="sxs-lookup"><span data-stu-id="6bc25-157">Requirements</span></span>

|<span data-ttu-id="6bc25-158">要求</span><span class="sxs-lookup"><span data-stu-id="6bc25-158">Requirement</span></span>| <span data-ttu-id="6bc25-159">值</span><span class="sxs-lookup"><span data-stu-id="6bc25-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="6bc25-160">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="6bc25-160">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6bc25-161">1.0</span><span class="sxs-lookup"><span data-stu-id="6bc25-161">1.0</span></span>|
|[<span data-ttu-id="6bc25-162">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="6bc25-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6bc25-163">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="6bc25-163">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="6bc25-164">EventType :String</span><span class="sxs-lookup"><span data-stu-id="6bc25-164">EventType :String</span></span>

<span data-ttu-id="6bc25-165">指定与事件处理程序相关联的事件。</span><span class="sxs-lookup"><span data-stu-id="6bc25-165">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="6bc25-166">类型：</span><span class="sxs-lookup"><span data-stu-id="6bc25-166">Type:</span></span>

*   <span data-ttu-id="6bc25-167">字符串</span><span class="sxs-lookup"><span data-stu-id="6bc25-167">String</span></span>

##### <a name="properties"></a><span data-ttu-id="6bc25-168">属性：</span><span class="sxs-lookup"><span data-stu-id="6bc25-168">Properties:</span></span>

| <span data-ttu-id="6bc25-169">名称</span><span class="sxs-lookup"><span data-stu-id="6bc25-169">Name</span></span> | <span data-ttu-id="6bc25-170">类型</span><span class="sxs-lookup"><span data-stu-id="6bc25-170">Type</span></span> | <span data-ttu-id="6bc25-171">说明</span><span class="sxs-lookup"><span data-stu-id="6bc25-171">Description</span></span> | <span data-ttu-id="6bc25-172">最低要求集</span><span class="sxs-lookup"><span data-stu-id="6bc25-172">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="6bc25-173">String</span><span class="sxs-lookup"><span data-stu-id="6bc25-173">String</span></span> | <span data-ttu-id="6bc25-174">已更改的日期或时间所选的约会或系列。</span><span class="sxs-lookup"><span data-stu-id="6bc25-174">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="6bc25-175">1.7</span><span class="sxs-lookup"><span data-stu-id="6bc25-175">1.7</span></span> |
|`ItemChanged`| <span data-ttu-id="6bc25-176">字符串</span><span class="sxs-lookup"><span data-stu-id="6bc25-176">String</span></span> | <span data-ttu-id="6bc25-177">选定的项已更改。</span><span class="sxs-lookup"><span data-stu-id="6bc25-177">The selected item has changed.</span></span> | <span data-ttu-id="6bc25-178">1.5</span><span class="sxs-lookup"><span data-stu-id="6bc25-178">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="6bc25-179">String</span><span class="sxs-lookup"><span data-stu-id="6bc25-179">String</span></span> | <span data-ttu-id="6bc25-180">已更改所选的项目或约会位置的收件人列表。</span><span class="sxs-lookup"><span data-stu-id="6bc25-180">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="6bc25-181">1.7</span><span class="sxs-lookup"><span data-stu-id="6bc25-181">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="6bc25-182">String</span><span class="sxs-lookup"><span data-stu-id="6bc25-182">String</span></span> | <span data-ttu-id="6bc25-183">所选系列的定期模式已更改。</span><span class="sxs-lookup"><span data-stu-id="6bc25-183">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="6bc25-184">1.7</span><span class="sxs-lookup"><span data-stu-id="6bc25-184">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="6bc25-185">要求</span><span class="sxs-lookup"><span data-stu-id="6bc25-185">Requirements</span></span>

|<span data-ttu-id="6bc25-186">要求</span><span class="sxs-lookup"><span data-stu-id="6bc25-186">Requirement</span></span>| <span data-ttu-id="6bc25-187">值</span><span class="sxs-lookup"><span data-stu-id="6bc25-187">Value</span></span>|
|---|---|
|[<span data-ttu-id="6bc25-188">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="6bc25-188">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6bc25-189">1.5</span><span class="sxs-lookup"><span data-stu-id="6bc25-189">1.5</span></span> |
|[<span data-ttu-id="6bc25-190">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="6bc25-190">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6bc25-191">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="6bc25-191">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="6bc25-192">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="6bc25-192">SourceProperty :String</span></span>

<span data-ttu-id="6bc25-193">指定由调用方法返回的数据源。</span><span class="sxs-lookup"><span data-stu-id="6bc25-193">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="6bc25-194">类型：</span><span class="sxs-lookup"><span data-stu-id="6bc25-194">Type:</span></span>

*   <span data-ttu-id="6bc25-195">字符串</span><span class="sxs-lookup"><span data-stu-id="6bc25-195">String</span></span>

##### <a name="properties"></a><span data-ttu-id="6bc25-196">属性：</span><span class="sxs-lookup"><span data-stu-id="6bc25-196">Properties:</span></span>

|<span data-ttu-id="6bc25-197">名称</span><span class="sxs-lookup"><span data-stu-id="6bc25-197">Name</span></span>| <span data-ttu-id="6bc25-198">类型</span><span class="sxs-lookup"><span data-stu-id="6bc25-198">Type</span></span>| <span data-ttu-id="6bc25-199">说明</span><span class="sxs-lookup"><span data-stu-id="6bc25-199">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="6bc25-200">字符串</span><span class="sxs-lookup"><span data-stu-id="6bc25-200">String</span></span>|<span data-ttu-id="6bc25-201">数据源来自邮件的正文。</span><span class="sxs-lookup"><span data-stu-id="6bc25-201">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="6bc25-202">String</span><span class="sxs-lookup"><span data-stu-id="6bc25-202">String</span></span>|<span data-ttu-id="6bc25-203">数据源来自邮件的主题。</span><span class="sxs-lookup"><span data-stu-id="6bc25-203">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6bc25-204">要求</span><span class="sxs-lookup"><span data-stu-id="6bc25-204">Requirements</span></span>

|<span data-ttu-id="6bc25-205">要求</span><span class="sxs-lookup"><span data-stu-id="6bc25-205">Requirement</span></span>| <span data-ttu-id="6bc25-206">值</span><span class="sxs-lookup"><span data-stu-id="6bc25-206">Value</span></span>|
|---|---|
|[<span data-ttu-id="6bc25-207">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="6bc25-207">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6bc25-208">1.0</span><span class="sxs-lookup"><span data-stu-id="6bc25-208">1.0</span></span>|
|[<span data-ttu-id="6bc25-209">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="6bc25-209">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6bc25-210">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="6bc25-210">Compose or read</span></span>|