
# <a name="userprofile"></a><span data-ttu-id="11336-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="11336-101">userProfile</span></span>

### <span data-ttu-id="11336-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="11336-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="11336-104">要求</span><span class="sxs-lookup"><span data-stu-id="11336-104">Requirements</span></span>

|<span data-ttu-id="11336-105">要求</span><span class="sxs-lookup"><span data-stu-id="11336-105">Requirement</span></span>| <span data-ttu-id="11336-106">值</span><span class="sxs-lookup"><span data-stu-id="11336-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="11336-107">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="11336-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="11336-108">1.0</span><span class="sxs-lookup"><span data-stu-id="11336-108">1.0</span></span>|
|[<span data-ttu-id="11336-109">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="11336-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="11336-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="11336-110">ReadItem</span></span>|
|[<span data-ttu-id="11336-111">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="11336-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="11336-112">Compose 或 Read</span><span class="sxs-lookup"><span data-stu-id="11336-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="11336-113">成员和方法</span><span class="sxs-lookup"><span data-stu-id="11336-113">Members and methods</span></span>

| <span data-ttu-id="11336-114">成员</span><span class="sxs-lookup"><span data-stu-id="11336-114">Member</span></span> | <span data-ttu-id="11336-115">类型</span><span class="sxs-lookup"><span data-stu-id="11336-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="11336-116">accountType</span><span class="sxs-lookup"><span data-stu-id="11336-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="11336-117">成员</span><span class="sxs-lookup"><span data-stu-id="11336-117">Member</span></span> |
| [<span data-ttu-id="11336-118">displayName</span><span class="sxs-lookup"><span data-stu-id="11336-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="11336-119">成员</span><span class="sxs-lookup"><span data-stu-id="11336-119">Member</span></span> |
| [<span data-ttu-id="11336-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="11336-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="11336-121">成员</span><span class="sxs-lookup"><span data-stu-id="11336-121">Member</span></span> |
| [<span data-ttu-id="11336-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="11336-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="11336-123">成员</span><span class="sxs-lookup"><span data-stu-id="11336-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="11336-124">Members</span><span class="sxs-lookup"><span data-stu-id="11336-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="11336-125">accountType： 字符串</span><span class="sxs-lookup"><span data-stu-id="11336-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="11336-126">此成员是当前支持中有仅 Outlook 2016 for Mac、 构建 16.9.1212 和更高版本。</span><span class="sxs-lookup"><span data-stu-id="11336-126">This member is currently only supported in Outlook 2016 for Mac, build 16.9.1212 and greater.</span></span>

<span data-ttu-id="11336-127">获取与邮箱关联的用户的帐户类型。</span><span class="sxs-lookup"><span data-stu-id="11336-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="11336-128">下表列出了可能的值。</span><span class="sxs-lookup"><span data-stu-id="11336-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="11336-129">值</span><span class="sxs-lookup"><span data-stu-id="11336-129">Value</span></span> | <span data-ttu-id="11336-130">说明</span><span class="sxs-lookup"><span data-stu-id="11336-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="11336-131">邮箱是内部部署 Exchange 服务器上。</span><span class="sxs-lookup"><span data-stu-id="11336-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="11336-132">邮箱相关联的 Gmail 帐户。</span><span class="sxs-lookup"><span data-stu-id="11336-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="11336-133">邮箱相关联的 Office 365 工作或学校帐户。</span><span class="sxs-lookup"><span data-stu-id="11336-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="11336-134">邮箱相关联的个人 Outlook.com 帐户。</span><span class="sxs-lookup"><span data-stu-id="11336-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="11336-135">类型:</span><span class="sxs-lookup"><span data-stu-id="11336-135">Type:</span></span>

*   <span data-ttu-id="11336-136">String</span><span class="sxs-lookup"><span data-stu-id="11336-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="11336-137">要求</span><span class="sxs-lookup"><span data-stu-id="11336-137">Requirements</span></span>

|<span data-ttu-id="11336-138">要求</span><span class="sxs-lookup"><span data-stu-id="11336-138">Requirement</span></span>| <span data-ttu-id="11336-139">值</span><span class="sxs-lookup"><span data-stu-id="11336-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="11336-140">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="11336-140">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="11336-141">1.6</span><span class="sxs-lookup"><span data-stu-id="11336-141">1.6</span></span> |
|[<span data-ttu-id="11336-142">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="11336-142">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="11336-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="11336-143">ReadItem</span></span>|
|[<span data-ttu-id="11336-144">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="11336-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="11336-145">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="11336-145">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="11336-146">示例</span><span class="sxs-lookup"><span data-stu-id="11336-146">Example</span></span>

```
console.log(Office.context.mailbox.userProfile.accountType);
```

####  <a name="displayname-string"></a><span data-ttu-id="11336-147">displayName :String</span><span class="sxs-lookup"><span data-stu-id="11336-147">displayName :String</span></span>

<span data-ttu-id="11336-148">获取用户的显示名称。</span><span class="sxs-lookup"><span data-stu-id="11336-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="11336-149">类型：</span><span class="sxs-lookup"><span data-stu-id="11336-149">Type:</span></span>

*   <span data-ttu-id="11336-150">String</span><span class="sxs-lookup"><span data-stu-id="11336-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="11336-151">要求</span><span class="sxs-lookup"><span data-stu-id="11336-151">Requirements</span></span>

|<span data-ttu-id="11336-152">要求</span><span class="sxs-lookup"><span data-stu-id="11336-152">Requirement</span></span>| <span data-ttu-id="11336-153">值</span><span class="sxs-lookup"><span data-stu-id="11336-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="11336-154">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="11336-154">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="11336-155">1.0</span><span class="sxs-lookup"><span data-stu-id="11336-155">1.0</span></span>|
|[<span data-ttu-id="11336-156">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="11336-156">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="11336-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="11336-157">ReadItem</span></span>|
|[<span data-ttu-id="11336-158">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="11336-158">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="11336-159">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="11336-159">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="11336-160">示例</span><span class="sxs-lookup"><span data-stu-id="11336-160">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="11336-161">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="11336-161">emailAddress :String</span></span>

<span data-ttu-id="11336-162">获取用户的 SMTP 电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="11336-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="11336-163">类型：</span><span class="sxs-lookup"><span data-stu-id="11336-163">Type:</span></span>

*   <span data-ttu-id="11336-164">String</span><span class="sxs-lookup"><span data-stu-id="11336-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="11336-165">要求</span><span class="sxs-lookup"><span data-stu-id="11336-165">Requirements</span></span>

|<span data-ttu-id="11336-166">要求</span><span class="sxs-lookup"><span data-stu-id="11336-166">Requirement</span></span>| <span data-ttu-id="11336-167">值</span><span class="sxs-lookup"><span data-stu-id="11336-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="11336-168">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="11336-168">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="11336-169">1.0</span><span class="sxs-lookup"><span data-stu-id="11336-169">1.0</span></span>|
|[<span data-ttu-id="11336-170">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="11336-170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="11336-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="11336-171">ReadItem</span></span>|
|[<span data-ttu-id="11336-172">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="11336-172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="11336-173">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="11336-173">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="11336-174">示例</span><span class="sxs-lookup"><span data-stu-id="11336-174">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="11336-175">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="11336-175">timeZone :String</span></span>

<span data-ttu-id="11336-176">获取用户的默认时区。</span><span class="sxs-lookup"><span data-stu-id="11336-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="11336-177">类型：</span><span class="sxs-lookup"><span data-stu-id="11336-177">Type:</span></span>

*   <span data-ttu-id="11336-178">String</span><span class="sxs-lookup"><span data-stu-id="11336-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="11336-179">要求</span><span class="sxs-lookup"><span data-stu-id="11336-179">Requirements</span></span>

|<span data-ttu-id="11336-180">要求</span><span class="sxs-lookup"><span data-stu-id="11336-180">Requirement</span></span>| <span data-ttu-id="11336-181">值</span><span class="sxs-lookup"><span data-stu-id="11336-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="11336-182">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="11336-182">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="11336-183">1.0</span><span class="sxs-lookup"><span data-stu-id="11336-183">1.0</span></span>|
|[<span data-ttu-id="11336-184">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="11336-184">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="11336-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="11336-185">ReadItem</span></span>|
|[<span data-ttu-id="11336-186">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="11336-186">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="11336-187">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="11336-187">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="11336-188">示例</span><span class="sxs-lookup"><span data-stu-id="11336-188">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```