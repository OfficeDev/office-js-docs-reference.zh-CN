
# <a name="userprofile"></a><span data-ttu-id="b6768-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="b6768-101">userProfile</span></span>

### <span data-ttu-id="b6768-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="b6768-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="b6768-104">要求</span><span class="sxs-lookup"><span data-stu-id="b6768-104">Requirements</span></span>

|<span data-ttu-id="b6768-105">要求</span><span class="sxs-lookup"><span data-stu-id="b6768-105">Requirement</span></span>| <span data-ttu-id="b6768-106">值</span><span class="sxs-lookup"><span data-stu-id="b6768-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="b6768-107">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b6768-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b6768-108">1.0</span><span class="sxs-lookup"><span data-stu-id="b6768-108">1.0</span></span>|
|[<span data-ttu-id="b6768-109">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="b6768-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b6768-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b6768-110">ReadItem</span></span>|
|[<span data-ttu-id="b6768-111">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b6768-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b6768-112">Compose 或 Read</span><span class="sxs-lookup"><span data-stu-id="b6768-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="b6768-113">成员和方法</span><span class="sxs-lookup"><span data-stu-id="b6768-113">Members and methods</span></span>

| <span data-ttu-id="b6768-114">成员</span><span class="sxs-lookup"><span data-stu-id="b6768-114">Member</span></span> | <span data-ttu-id="b6768-115">类型</span><span class="sxs-lookup"><span data-stu-id="b6768-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="b6768-116">accountType</span><span class="sxs-lookup"><span data-stu-id="b6768-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="b6768-117">成员</span><span class="sxs-lookup"><span data-stu-id="b6768-117">Member</span></span> |
| [<span data-ttu-id="b6768-118">displayName</span><span class="sxs-lookup"><span data-stu-id="b6768-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="b6768-119">成员</span><span class="sxs-lookup"><span data-stu-id="b6768-119">Member</span></span> |
| [<span data-ttu-id="b6768-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="b6768-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="b6768-121">成员</span><span class="sxs-lookup"><span data-stu-id="b6768-121">Member</span></span> |
| [<span data-ttu-id="b6768-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="b6768-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="b6768-123">成员</span><span class="sxs-lookup"><span data-stu-id="b6768-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="b6768-124">Members</span><span class="sxs-lookup"><span data-stu-id="b6768-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="b6768-125">accountType： 字符串</span><span class="sxs-lookup"><span data-stu-id="b6768-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="b6768-126">此成员是当前支持中有仅 Outlook 2016 for Mac、 构建 16.9.1212 和更高版本。</span><span class="sxs-lookup"><span data-stu-id="b6768-126">This member is currently only supported in Outlook 2016 for Mac, build 16.9.1212 and greater.</span></span>

<span data-ttu-id="b6768-127">获取与邮箱关联的用户的帐户类型。</span><span class="sxs-lookup"><span data-stu-id="b6768-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="b6768-128">下表列出了可能的值。</span><span class="sxs-lookup"><span data-stu-id="b6768-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="b6768-129">值</span><span class="sxs-lookup"><span data-stu-id="b6768-129">Value</span></span> | <span data-ttu-id="b6768-130">说明</span><span class="sxs-lookup"><span data-stu-id="b6768-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="b6768-131">邮箱是内部部署 Exchange 服务器上。</span><span class="sxs-lookup"><span data-stu-id="b6768-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="b6768-132">邮箱相关联的 Gmail 帐户。</span><span class="sxs-lookup"><span data-stu-id="b6768-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="b6768-133">邮箱相关联的 Office 365 工作或学校帐户。</span><span class="sxs-lookup"><span data-stu-id="b6768-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="b6768-134">邮箱相关联的个人 Outlook.com 帐户。</span><span class="sxs-lookup"><span data-stu-id="b6768-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="b6768-135">类型:</span><span class="sxs-lookup"><span data-stu-id="b6768-135">Type:</span></span>

*   <span data-ttu-id="b6768-136">String</span><span class="sxs-lookup"><span data-stu-id="b6768-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b6768-137">要求</span><span class="sxs-lookup"><span data-stu-id="b6768-137">Requirements</span></span>

|<span data-ttu-id="b6768-138">要求</span><span class="sxs-lookup"><span data-stu-id="b6768-138">Requirement</span></span>| <span data-ttu-id="b6768-139">值</span><span class="sxs-lookup"><span data-stu-id="b6768-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="b6768-140">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b6768-140">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b6768-141">1.6</span><span class="sxs-lookup"><span data-stu-id="b6768-141">1.6</span></span> |
|[<span data-ttu-id="b6768-142">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="b6768-142">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b6768-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b6768-143">ReadItem</span></span>|
|[<span data-ttu-id="b6768-144">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b6768-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b6768-145">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b6768-145">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b6768-146">示例</span><span class="sxs-lookup"><span data-stu-id="b6768-146">Example</span></span>

```
console.log(Office.context.mailbox.userProfile.accountType);
```

####  <a name="displayname-string"></a><span data-ttu-id="b6768-147">displayName :String</span><span class="sxs-lookup"><span data-stu-id="b6768-147">displayName :String</span></span>

<span data-ttu-id="b6768-148">获取用户的显示名称。</span><span class="sxs-lookup"><span data-stu-id="b6768-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="b6768-149">类型：</span><span class="sxs-lookup"><span data-stu-id="b6768-149">Type:</span></span>

*   <span data-ttu-id="b6768-150">String</span><span class="sxs-lookup"><span data-stu-id="b6768-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b6768-151">要求</span><span class="sxs-lookup"><span data-stu-id="b6768-151">Requirements</span></span>

|<span data-ttu-id="b6768-152">要求</span><span class="sxs-lookup"><span data-stu-id="b6768-152">Requirement</span></span>| <span data-ttu-id="b6768-153">值</span><span class="sxs-lookup"><span data-stu-id="b6768-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="b6768-154">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b6768-154">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b6768-155">1.0</span><span class="sxs-lookup"><span data-stu-id="b6768-155">1.0</span></span>|
|[<span data-ttu-id="b6768-156">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="b6768-156">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b6768-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b6768-157">ReadItem</span></span>|
|[<span data-ttu-id="b6768-158">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b6768-158">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b6768-159">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b6768-159">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b6768-160">示例</span><span class="sxs-lookup"><span data-stu-id="b6768-160">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="b6768-161">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="b6768-161">emailAddress :String</span></span>

<span data-ttu-id="b6768-162">获取用户的 SMTP 电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="b6768-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="b6768-163">类型：</span><span class="sxs-lookup"><span data-stu-id="b6768-163">Type:</span></span>

*   <span data-ttu-id="b6768-164">String</span><span class="sxs-lookup"><span data-stu-id="b6768-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b6768-165">要求</span><span class="sxs-lookup"><span data-stu-id="b6768-165">Requirements</span></span>

|<span data-ttu-id="b6768-166">要求</span><span class="sxs-lookup"><span data-stu-id="b6768-166">Requirement</span></span>| <span data-ttu-id="b6768-167">值</span><span class="sxs-lookup"><span data-stu-id="b6768-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="b6768-168">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b6768-168">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b6768-169">1.0</span><span class="sxs-lookup"><span data-stu-id="b6768-169">1.0</span></span>|
|[<span data-ttu-id="b6768-170">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="b6768-170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b6768-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b6768-171">ReadItem</span></span>|
|[<span data-ttu-id="b6768-172">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b6768-172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b6768-173">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b6768-173">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b6768-174">示例</span><span class="sxs-lookup"><span data-stu-id="b6768-174">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="b6768-175">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="b6768-175">timeZone :String</span></span>

<span data-ttu-id="b6768-176">获取用户的默认时区。</span><span class="sxs-lookup"><span data-stu-id="b6768-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="b6768-177">类型：</span><span class="sxs-lookup"><span data-stu-id="b6768-177">Type:</span></span>

*   <span data-ttu-id="b6768-178">String</span><span class="sxs-lookup"><span data-stu-id="b6768-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b6768-179">要求</span><span class="sxs-lookup"><span data-stu-id="b6768-179">Requirements</span></span>

|<span data-ttu-id="b6768-180">要求</span><span class="sxs-lookup"><span data-stu-id="b6768-180">Requirement</span></span>| <span data-ttu-id="b6768-181">值</span><span class="sxs-lookup"><span data-stu-id="b6768-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="b6768-182">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b6768-182">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b6768-183">1.0</span><span class="sxs-lookup"><span data-stu-id="b6768-183">1.0</span></span>|
|[<span data-ttu-id="b6768-184">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="b6768-184">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b6768-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b6768-185">ReadItem</span></span>|
|[<span data-ttu-id="b6768-186">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b6768-186">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b6768-187">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b6768-187">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b6768-188">示例</span><span class="sxs-lookup"><span data-stu-id="b6768-188">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```