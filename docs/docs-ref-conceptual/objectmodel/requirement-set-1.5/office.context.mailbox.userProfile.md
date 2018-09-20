# <a name="userprofile"></a><span data-ttu-id="5997c-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="5997c-101">userProfile</span></span>

### <span data-ttu-id="5997c-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="5997c-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="5997c-104">要求</span><span class="sxs-lookup"><span data-stu-id="5997c-104">Requirements</span></span>

|<span data-ttu-id="5997c-105">要求</span><span class="sxs-lookup"><span data-stu-id="5997c-105">Requirement</span></span>| <span data-ttu-id="5997c-106">值</span><span class="sxs-lookup"><span data-stu-id="5997c-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="5997c-107">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="5997c-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5997c-108">1.0</span><span class="sxs-lookup"><span data-stu-id="5997c-108">1.0</span></span>|
|[<span data-ttu-id="5997c-109">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="5997c-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5997c-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5997c-110">ReadItem</span></span>|
|[<span data-ttu-id="5997c-111">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="5997c-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5997c-112">Compose 或 Read</span><span class="sxs-lookup"><span data-stu-id="5997c-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="5997c-113">成员和方法</span><span class="sxs-lookup"><span data-stu-id="5997c-113">Members and methods</span></span>

| <span data-ttu-id="5997c-114">成员</span><span class="sxs-lookup"><span data-stu-id="5997c-114">Member</span></span> | <span data-ttu-id="5997c-115">类型</span><span class="sxs-lookup"><span data-stu-id="5997c-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="5997c-116">displayName</span><span class="sxs-lookup"><span data-stu-id="5997c-116">displayName</span></span>](#displayname-string) | <span data-ttu-id="5997c-117">成员</span><span class="sxs-lookup"><span data-stu-id="5997c-117">Member</span></span> |
| [<span data-ttu-id="5997c-118">emailAddress</span><span class="sxs-lookup"><span data-stu-id="5997c-118">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="5997c-119">成员</span><span class="sxs-lookup"><span data-stu-id="5997c-119">Member</span></span> |
| [<span data-ttu-id="5997c-120">timeZone</span><span class="sxs-lookup"><span data-stu-id="5997c-120">timeZone</span></span>](#timezone-string) | <span data-ttu-id="5997c-121">成员</span><span class="sxs-lookup"><span data-stu-id="5997c-121">Member</span></span> |

### <a name="members"></a><span data-ttu-id="5997c-122">成员</span><span class="sxs-lookup"><span data-stu-id="5997c-122">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="5997c-123">displayName :String</span><span class="sxs-lookup"><span data-stu-id="5997c-123">displayName :String</span></span>

<span data-ttu-id="5997c-124">获取用户的显示名称。</span><span class="sxs-lookup"><span data-stu-id="5997c-124">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="5997c-125">类型：</span><span class="sxs-lookup"><span data-stu-id="5997c-125">Type:</span></span>

*   <span data-ttu-id="5997c-126">String</span><span class="sxs-lookup"><span data-stu-id="5997c-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="5997c-127">要求</span><span class="sxs-lookup"><span data-stu-id="5997c-127">Requirements</span></span>

|<span data-ttu-id="5997c-128">要求</span><span class="sxs-lookup"><span data-stu-id="5997c-128">Requirement</span></span>| <span data-ttu-id="5997c-129">值</span><span class="sxs-lookup"><span data-stu-id="5997c-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="5997c-130">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="5997c-130">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5997c-131">1.0</span><span class="sxs-lookup"><span data-stu-id="5997c-131">1.0</span></span>|
|[<span data-ttu-id="5997c-132">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="5997c-132">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5997c-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5997c-133">ReadItem</span></span>|
|[<span data-ttu-id="5997c-134">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="5997c-134">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5997c-135">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="5997c-135">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="5997c-136">示例</span><span class="sxs-lookup"><span data-stu-id="5997c-136">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="5997c-137">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="5997c-137">emailAddress :String</span></span>

<span data-ttu-id="5997c-138">获取用户的 SMTP 电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="5997c-138">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="5997c-139">类型：</span><span class="sxs-lookup"><span data-stu-id="5997c-139">Type:</span></span>

*   <span data-ttu-id="5997c-140">String</span><span class="sxs-lookup"><span data-stu-id="5997c-140">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="5997c-141">要求</span><span class="sxs-lookup"><span data-stu-id="5997c-141">Requirements</span></span>

|<span data-ttu-id="5997c-142">要求</span><span class="sxs-lookup"><span data-stu-id="5997c-142">Requirement</span></span>| <span data-ttu-id="5997c-143">值</span><span class="sxs-lookup"><span data-stu-id="5997c-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="5997c-144">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="5997c-144">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5997c-145">1.0</span><span class="sxs-lookup"><span data-stu-id="5997c-145">1.0</span></span>|
|[<span data-ttu-id="5997c-146">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="5997c-146">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5997c-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5997c-147">ReadItem</span></span>|
|[<span data-ttu-id="5997c-148">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="5997c-148">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5997c-149">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="5997c-149">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="5997c-150">示例</span><span class="sxs-lookup"><span data-stu-id="5997c-150">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="5997c-151">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="5997c-151">timeZone :String</span></span>

<span data-ttu-id="5997c-152">获取用户的默认时区。</span><span class="sxs-lookup"><span data-stu-id="5997c-152">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="5997c-153">类型：</span><span class="sxs-lookup"><span data-stu-id="5997c-153">Type:</span></span>

*   <span data-ttu-id="5997c-154">String</span><span class="sxs-lookup"><span data-stu-id="5997c-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="5997c-155">要求</span><span class="sxs-lookup"><span data-stu-id="5997c-155">Requirements</span></span>

|<span data-ttu-id="5997c-156">要求</span><span class="sxs-lookup"><span data-stu-id="5997c-156">Requirement</span></span>| <span data-ttu-id="5997c-157">值</span><span class="sxs-lookup"><span data-stu-id="5997c-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="5997c-158">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="5997c-158">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5997c-159">1.0</span><span class="sxs-lookup"><span data-stu-id="5997c-159">1.0</span></span>|
|[<span data-ttu-id="5997c-160">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="5997c-160">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5997c-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5997c-161">ReadItem</span></span>|
|[<span data-ttu-id="5997c-162">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="5997c-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5997c-163">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="5997c-163">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="5997c-164">示例</span><span class="sxs-lookup"><span data-stu-id="5997c-164">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```