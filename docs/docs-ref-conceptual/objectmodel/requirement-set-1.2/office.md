 

# <a name="office"></a><span data-ttu-id="87f93-101">Office</span><span class="sxs-lookup"><span data-stu-id="87f93-101">Office</span></span>

<span data-ttu-id="87f93-p101">该 Office 命名空间提供所有 Office 应用中的外接程序所使用的共享接口。此列表仅记录 Outlook 外接程序所使用的接口。有关 Office 命名空间的完整列表，请参阅[共享 API](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="87f93-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="87f93-104">Requirements</span><span class="sxs-lookup"><span data-stu-id="87f93-104">Requirements</span></span>

|<span data-ttu-id="87f93-105">要求</span><span class="sxs-lookup"><span data-stu-id="87f93-105">Requirement</span></span>| <span data-ttu-id="87f93-106">值</span><span class="sxs-lookup"><span data-stu-id="87f93-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="87f93-107">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="87f93-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="87f93-108">1.0</span><span class="sxs-lookup"><span data-stu-id="87f93-108">1.0</span></span>|
|[<span data-ttu-id="87f93-109">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="87f93-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="87f93-110">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="87f93-110">Compose or read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="87f93-111">命名空间</span><span class="sxs-lookup"><span data-stu-id="87f93-111">Namespaces</span></span>

<span data-ttu-id="87f93-112">[上下文](office.context.md)： 用于 Outlook 外接程序 API 提供共享的 Office 加载项 API 上下文命名空间中的接口。</span><span class="sxs-lookup"><span data-stu-id="87f93-112">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="87f93-113">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype)：包括 ItemType、EntityType、AttachmentType、RecipientType、ResponseType 和 ItemNotificationMessageType 枚举。</span><span class="sxs-lookup"><span data-stu-id="87f93-113">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="87f93-114">成员</span><span class="sxs-lookup"><span data-stu-id="87f93-114">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="87f93-115">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="87f93-115">AsyncResultStatus :String</span></span>

<span data-ttu-id="87f93-116">指定异步调用的结果。</span><span class="sxs-lookup"><span data-stu-id="87f93-116">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="87f93-117">类型：</span><span class="sxs-lookup"><span data-stu-id="87f93-117">Type:</span></span>

*   <span data-ttu-id="87f93-118">字符串</span><span class="sxs-lookup"><span data-stu-id="87f93-118">String</span></span>

##### <a name="properties"></a><span data-ttu-id="87f93-119">属性：</span><span class="sxs-lookup"><span data-stu-id="87f93-119">Properties:</span></span>

|<span data-ttu-id="87f93-120">名称</span><span class="sxs-lookup"><span data-stu-id="87f93-120">Name</span></span>| <span data-ttu-id="87f93-121">类型</span><span class="sxs-lookup"><span data-stu-id="87f93-121">Type</span></span>| <span data-ttu-id="87f93-122">说明</span><span class="sxs-lookup"><span data-stu-id="87f93-122">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="87f93-123">String</span><span class="sxs-lookup"><span data-stu-id="87f93-123">String</span></span>|<span data-ttu-id="87f93-124">调用成功。</span><span class="sxs-lookup"><span data-stu-id="87f93-124">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="87f93-125">字符串</span><span class="sxs-lookup"><span data-stu-id="87f93-125">String</span></span>|<span data-ttu-id="87f93-126">调用失败。</span><span class="sxs-lookup"><span data-stu-id="87f93-126">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="87f93-127">要求</span><span class="sxs-lookup"><span data-stu-id="87f93-127">Requirements</span></span>

|<span data-ttu-id="87f93-128">要求</span><span class="sxs-lookup"><span data-stu-id="87f93-128">Requirement</span></span>| <span data-ttu-id="87f93-129">值</span><span class="sxs-lookup"><span data-stu-id="87f93-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="87f93-130">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="87f93-130">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="87f93-131">1.0</span><span class="sxs-lookup"><span data-stu-id="87f93-131">1.0</span></span>|
|[<span data-ttu-id="87f93-132">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="87f93-132">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="87f93-133">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="87f93-133">Compose or read</span></span>|
####  <a name="coerciontype-string"></a><span data-ttu-id="87f93-134">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="87f93-134">CoercionType :String</span></span>

<span data-ttu-id="87f93-135">指定如何强制由调用方法返回或设置的数据。</span><span class="sxs-lookup"><span data-stu-id="87f93-135">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="87f93-136">类型：</span><span class="sxs-lookup"><span data-stu-id="87f93-136">Type:</span></span>

*   <span data-ttu-id="87f93-137">字符串</span><span class="sxs-lookup"><span data-stu-id="87f93-137">String</span></span>

##### <a name="properties"></a><span data-ttu-id="87f93-138">属性：</span><span class="sxs-lookup"><span data-stu-id="87f93-138">Properties:</span></span>

|<span data-ttu-id="87f93-139">名称</span><span class="sxs-lookup"><span data-stu-id="87f93-139">Name</span></span>| <span data-ttu-id="87f93-140">类型</span><span class="sxs-lookup"><span data-stu-id="87f93-140">Type</span></span>| <span data-ttu-id="87f93-141">说明</span><span class="sxs-lookup"><span data-stu-id="87f93-141">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="87f93-142">String</span><span class="sxs-lookup"><span data-stu-id="87f93-142">String</span></span>|<span data-ttu-id="87f93-143">请求以 HTML 格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="87f93-143">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="87f93-144">字符串</span><span class="sxs-lookup"><span data-stu-id="87f93-144">String</span></span>|<span data-ttu-id="87f93-145">请求以文本格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="87f93-145">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="87f93-146">要求</span><span class="sxs-lookup"><span data-stu-id="87f93-146">Requirements</span></span>

|<span data-ttu-id="87f93-147">要求</span><span class="sxs-lookup"><span data-stu-id="87f93-147">Requirement</span></span>| <span data-ttu-id="87f93-148">值</span><span class="sxs-lookup"><span data-stu-id="87f93-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="87f93-149">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="87f93-149">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="87f93-150">1.0</span><span class="sxs-lookup"><span data-stu-id="87f93-150">1.0</span></span>|
|[<span data-ttu-id="87f93-151">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="87f93-151">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="87f93-152">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="87f93-152">Compose or read</span></span>|
####  <a name="sourceproperty-string"></a><span data-ttu-id="87f93-153">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="87f93-153">SourceProperty :String</span></span>

<span data-ttu-id="87f93-154">指定由调用方法返回的数据源。</span><span class="sxs-lookup"><span data-stu-id="87f93-154">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="87f93-155">类型：</span><span class="sxs-lookup"><span data-stu-id="87f93-155">Type:</span></span>

*   <span data-ttu-id="87f93-156">字符串</span><span class="sxs-lookup"><span data-stu-id="87f93-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="87f93-157">属性：</span><span class="sxs-lookup"><span data-stu-id="87f93-157">Properties:</span></span>

|<span data-ttu-id="87f93-158">名称</span><span class="sxs-lookup"><span data-stu-id="87f93-158">Name</span></span>| <span data-ttu-id="87f93-159">类型</span><span class="sxs-lookup"><span data-stu-id="87f93-159">Type</span></span>| <span data-ttu-id="87f93-160">说明</span><span class="sxs-lookup"><span data-stu-id="87f93-160">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="87f93-161">字符串</span><span class="sxs-lookup"><span data-stu-id="87f93-161">String</span></span>|<span data-ttu-id="87f93-162">数据源来自邮件的正文。</span><span class="sxs-lookup"><span data-stu-id="87f93-162">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="87f93-163">String</span><span class="sxs-lookup"><span data-stu-id="87f93-163">String</span></span>|<span data-ttu-id="87f93-164">数据源来自邮件的主题。</span><span class="sxs-lookup"><span data-stu-id="87f93-164">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="87f93-165">要求</span><span class="sxs-lookup"><span data-stu-id="87f93-165">Requirements</span></span>

|<span data-ttu-id="87f93-166">要求</span><span class="sxs-lookup"><span data-stu-id="87f93-166">Requirement</span></span>| <span data-ttu-id="87f93-167">值</span><span class="sxs-lookup"><span data-stu-id="87f93-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="87f93-168">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="87f93-168">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="87f93-169">1.0</span><span class="sxs-lookup"><span data-stu-id="87f93-169">1.0</span></span>|
|[<span data-ttu-id="87f93-170">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="87f93-170">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="87f93-171">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="87f93-171">Compose or read</span></span>|