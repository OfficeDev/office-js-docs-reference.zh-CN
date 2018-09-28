# <a name="office"></a><span data-ttu-id="4fc39-101">Office</span><span class="sxs-lookup"><span data-stu-id="4fc39-101">Office</span></span>

<span data-ttu-id="4fc39-p101">该 Office 命名空间提供所有 Office 应用中的外接程序所使用的共享接口。此列表仅记录 Outlook 外接程序所使用的接口。有关 Office 命名空间的完整列表，请参阅[共享 API](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="4fc39-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="4fc39-104">Requirements</span><span class="sxs-lookup"><span data-stu-id="4fc39-104">Requirements</span></span>

|<span data-ttu-id="4fc39-105">要求</span><span class="sxs-lookup"><span data-stu-id="4fc39-105">Requirement</span></span>| <span data-ttu-id="4fc39-106">值</span><span class="sxs-lookup"><span data-stu-id="4fc39-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="4fc39-107">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4fc39-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4fc39-108">1.0</span><span class="sxs-lookup"><span data-stu-id="4fc39-108">1.0</span></span>|
|[<span data-ttu-id="4fc39-109">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4fc39-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4fc39-110">Compose 或 Read</span><span class="sxs-lookup"><span data-stu-id="4fc39-110">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="4fc39-111">成员和方法</span><span class="sxs-lookup"><span data-stu-id="4fc39-111">Members and methods</span></span>

| <span data-ttu-id="4fc39-112">成员</span><span class="sxs-lookup"><span data-stu-id="4fc39-112">Member</span></span> | <span data-ttu-id="4fc39-113">类型</span><span class="sxs-lookup"><span data-stu-id="4fc39-113">Type</span></span> |
|--------|------|
| [<span data-ttu-id="4fc39-114">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="4fc39-114">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="4fc39-115">成员</span><span class="sxs-lookup"><span data-stu-id="4fc39-115">Member</span></span> |
| [<span data-ttu-id="4fc39-116">CoercionType</span><span class="sxs-lookup"><span data-stu-id="4fc39-116">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="4fc39-117">成员</span><span class="sxs-lookup"><span data-stu-id="4fc39-117">Member</span></span> |
| [<span data-ttu-id="4fc39-118">EventType</span><span class="sxs-lookup"><span data-stu-id="4fc39-118">EventType</span></span>](#eventtype-string) | <span data-ttu-id="4fc39-119">成员</span><span class="sxs-lookup"><span data-stu-id="4fc39-119">Member</span></span> |
| [<span data-ttu-id="4fc39-120">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="4fc39-120">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="4fc39-121">成员</span><span class="sxs-lookup"><span data-stu-id="4fc39-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="4fc39-122">命名空间</span><span class="sxs-lookup"><span data-stu-id="4fc39-122">Namespaces</span></span>

<span data-ttu-id="4fc39-123">[上下文](office.context.md)： 用于 Outlook 外接程序 API 提供共享的 Office 加载项 API 上下文命名空间中的接口。</span><span class="sxs-lookup"><span data-stu-id="4fc39-123">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="4fc39-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype)：包括 ItemType、EntityType、AttachmentType、RecipientType、ResponseType 和 ItemNotificationMessageType 枚举。</span><span class="sxs-lookup"><span data-stu-id="4fc39-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="4fc39-125">成员</span><span class="sxs-lookup"><span data-stu-id="4fc39-125">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="4fc39-126">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="4fc39-126">AsyncResultStatus :String</span></span>

<span data-ttu-id="4fc39-127">指定异步调用的结果。</span><span class="sxs-lookup"><span data-stu-id="4fc39-127">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="4fc39-128">类型：</span><span class="sxs-lookup"><span data-stu-id="4fc39-128">Type:</span></span>

*   <span data-ttu-id="4fc39-129">字符串</span><span class="sxs-lookup"><span data-stu-id="4fc39-129">String</span></span>

##### <a name="properties"></a><span data-ttu-id="4fc39-130">属性：</span><span class="sxs-lookup"><span data-stu-id="4fc39-130">Properties:</span></span>

|<span data-ttu-id="4fc39-131">名称</span><span class="sxs-lookup"><span data-stu-id="4fc39-131">Name</span></span>| <span data-ttu-id="4fc39-132">类型</span><span class="sxs-lookup"><span data-stu-id="4fc39-132">Type</span></span>| <span data-ttu-id="4fc39-133">说明</span><span class="sxs-lookup"><span data-stu-id="4fc39-133">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="4fc39-134">String</span><span class="sxs-lookup"><span data-stu-id="4fc39-134">String</span></span>|<span data-ttu-id="4fc39-135">调用成功。</span><span class="sxs-lookup"><span data-stu-id="4fc39-135">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="4fc39-136">字符串</span><span class="sxs-lookup"><span data-stu-id="4fc39-136">String</span></span>|<span data-ttu-id="4fc39-137">调用失败。</span><span class="sxs-lookup"><span data-stu-id="4fc39-137">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4fc39-138">要求</span><span class="sxs-lookup"><span data-stu-id="4fc39-138">Requirements</span></span>

|<span data-ttu-id="4fc39-139">要求</span><span class="sxs-lookup"><span data-stu-id="4fc39-139">Requirement</span></span>| <span data-ttu-id="4fc39-140">值</span><span class="sxs-lookup"><span data-stu-id="4fc39-140">Value</span></span>|
|---|---|
|[<span data-ttu-id="4fc39-141">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4fc39-141">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4fc39-142">1.0</span><span class="sxs-lookup"><span data-stu-id="4fc39-142">1.0</span></span>|
|[<span data-ttu-id="4fc39-143">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4fc39-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4fc39-144">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4fc39-144">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="4fc39-145">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="4fc39-145">CoercionType :String</span></span>

<span data-ttu-id="4fc39-146">指定如何强制由调用方法返回或设置的数据。</span><span class="sxs-lookup"><span data-stu-id="4fc39-146">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="4fc39-147">类型：</span><span class="sxs-lookup"><span data-stu-id="4fc39-147">Type:</span></span>

*   <span data-ttu-id="4fc39-148">字符串</span><span class="sxs-lookup"><span data-stu-id="4fc39-148">String</span></span>

##### <a name="properties"></a><span data-ttu-id="4fc39-149">属性：</span><span class="sxs-lookup"><span data-stu-id="4fc39-149">Properties:</span></span>

|<span data-ttu-id="4fc39-150">名称</span><span class="sxs-lookup"><span data-stu-id="4fc39-150">Name</span></span>| <span data-ttu-id="4fc39-151">类型</span><span class="sxs-lookup"><span data-stu-id="4fc39-151">Type</span></span>| <span data-ttu-id="4fc39-152">说明</span><span class="sxs-lookup"><span data-stu-id="4fc39-152">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="4fc39-153">String</span><span class="sxs-lookup"><span data-stu-id="4fc39-153">String</span></span>|<span data-ttu-id="4fc39-154">请求以 HTML 格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="4fc39-154">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="4fc39-155">字符串</span><span class="sxs-lookup"><span data-stu-id="4fc39-155">String</span></span>|<span data-ttu-id="4fc39-156">请求以文本格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="4fc39-156">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4fc39-157">要求</span><span class="sxs-lookup"><span data-stu-id="4fc39-157">Requirements</span></span>

|<span data-ttu-id="4fc39-158">要求</span><span class="sxs-lookup"><span data-stu-id="4fc39-158">Requirement</span></span>| <span data-ttu-id="4fc39-159">值</span><span class="sxs-lookup"><span data-stu-id="4fc39-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="4fc39-160">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4fc39-160">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4fc39-161">1.0</span><span class="sxs-lookup"><span data-stu-id="4fc39-161">1.0</span></span>|
|[<span data-ttu-id="4fc39-162">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4fc39-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4fc39-163">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4fc39-163">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="4fc39-164">EventType :String</span><span class="sxs-lookup"><span data-stu-id="4fc39-164">EventType :String</span></span>

<span data-ttu-id="4fc39-165">指定与事件处理程序相关联的事件。</span><span class="sxs-lookup"><span data-stu-id="4fc39-165">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="4fc39-166">类型：</span><span class="sxs-lookup"><span data-stu-id="4fc39-166">Type:</span></span>

*   <span data-ttu-id="4fc39-167">字符串</span><span class="sxs-lookup"><span data-stu-id="4fc39-167">String</span></span>

##### <a name="properties"></a><span data-ttu-id="4fc39-168">属性：</span><span class="sxs-lookup"><span data-stu-id="4fc39-168">Properties:</span></span>

| <span data-ttu-id="4fc39-169">名称</span><span class="sxs-lookup"><span data-stu-id="4fc39-169">Name</span></span> | <span data-ttu-id="4fc39-170">类型</span><span class="sxs-lookup"><span data-stu-id="4fc39-170">Type</span></span> | <span data-ttu-id="4fc39-171">说明</span><span class="sxs-lookup"><span data-stu-id="4fc39-171">Description</span></span> |
|---|---|---|
|`ItemChanged`| <span data-ttu-id="4fc39-172">字符串</span><span class="sxs-lookup"><span data-stu-id="4fc39-172">String</span></span> | <span data-ttu-id="4fc39-173">选定的项已更改。</span><span class="sxs-lookup"><span data-stu-id="4fc39-173">The selected item has changed.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="4fc39-174">要求</span><span class="sxs-lookup"><span data-stu-id="4fc39-174">Requirements</span></span>

|<span data-ttu-id="4fc39-175">要求</span><span class="sxs-lookup"><span data-stu-id="4fc39-175">Requirement</span></span>| <span data-ttu-id="4fc39-176">值</span><span class="sxs-lookup"><span data-stu-id="4fc39-176">Value</span></span>|
|---|---|
|[<span data-ttu-id="4fc39-177">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4fc39-177">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4fc39-178">1.5</span><span class="sxs-lookup"><span data-stu-id="4fc39-178">1.5</span></span> |
|[<span data-ttu-id="4fc39-179">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4fc39-179">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4fc39-180">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4fc39-180">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="4fc39-181">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="4fc39-181">SourceProperty :String</span></span>

<span data-ttu-id="4fc39-182">指定由调用方法返回的数据源。</span><span class="sxs-lookup"><span data-stu-id="4fc39-182">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="4fc39-183">类型：</span><span class="sxs-lookup"><span data-stu-id="4fc39-183">Type:</span></span>

*   <span data-ttu-id="4fc39-184">字符串</span><span class="sxs-lookup"><span data-stu-id="4fc39-184">String</span></span>

##### <a name="properties"></a><span data-ttu-id="4fc39-185">属性：</span><span class="sxs-lookup"><span data-stu-id="4fc39-185">Properties:</span></span>

|<span data-ttu-id="4fc39-186">名称</span><span class="sxs-lookup"><span data-stu-id="4fc39-186">Name</span></span>| <span data-ttu-id="4fc39-187">类型</span><span class="sxs-lookup"><span data-stu-id="4fc39-187">Type</span></span>| <span data-ttu-id="4fc39-188">说明</span><span class="sxs-lookup"><span data-stu-id="4fc39-188">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="4fc39-189">字符串</span><span class="sxs-lookup"><span data-stu-id="4fc39-189">String</span></span>|<span data-ttu-id="4fc39-190">数据源来自邮件的正文。</span><span class="sxs-lookup"><span data-stu-id="4fc39-190">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="4fc39-191">String</span><span class="sxs-lookup"><span data-stu-id="4fc39-191">String</span></span>|<span data-ttu-id="4fc39-192">数据源来自邮件的主题。</span><span class="sxs-lookup"><span data-stu-id="4fc39-192">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4fc39-193">要求</span><span class="sxs-lookup"><span data-stu-id="4fc39-193">Requirements</span></span>

|<span data-ttu-id="4fc39-194">要求</span><span class="sxs-lookup"><span data-stu-id="4fc39-194">Requirement</span></span>| <span data-ttu-id="4fc39-195">值</span><span class="sxs-lookup"><span data-stu-id="4fc39-195">Value</span></span>|
|---|---|
|[<span data-ttu-id="4fc39-196">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4fc39-196">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4fc39-197">1.0</span><span class="sxs-lookup"><span data-stu-id="4fc39-197">1.0</span></span>|
|[<span data-ttu-id="4fc39-198">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4fc39-198">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4fc39-199">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4fc39-199">Compose or read</span></span>|