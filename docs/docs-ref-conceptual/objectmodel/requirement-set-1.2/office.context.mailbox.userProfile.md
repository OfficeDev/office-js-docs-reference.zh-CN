
# <a name="userprofile"></a><span data-ttu-id="85836-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="85836-101">userProfile</span></span>

### <span data-ttu-id="85836-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="85836-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="85836-104">要求</span><span class="sxs-lookup"><span data-stu-id="85836-104">Requirements</span></span>

|<span data-ttu-id="85836-105">要求</span><span class="sxs-lookup"><span data-stu-id="85836-105">Requirement</span></span>| <span data-ttu-id="85836-106">值</span><span class="sxs-lookup"><span data-stu-id="85836-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="85836-107">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="85836-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="85836-108">1.0</span><span class="sxs-lookup"><span data-stu-id="85836-108">1.0</span></span>|
|[<span data-ttu-id="85836-109">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="85836-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="85836-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="85836-110">ReadItem</span></span>|
|[<span data-ttu-id="85836-111">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="85836-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="85836-112">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="85836-112">Compose or read</span></span>|

### <a name="members"></a><span data-ttu-id="85836-113">成员</span><span class="sxs-lookup"><span data-stu-id="85836-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="85836-114">displayName :String</span><span class="sxs-lookup"><span data-stu-id="85836-114">displayName :String</span></span>

<span data-ttu-id="85836-115">获取用户的显示名称。</span><span class="sxs-lookup"><span data-stu-id="85836-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="85836-116">类型：</span><span class="sxs-lookup"><span data-stu-id="85836-116">Type:</span></span>

*   <span data-ttu-id="85836-117">String</span><span class="sxs-lookup"><span data-stu-id="85836-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="85836-118">要求</span><span class="sxs-lookup"><span data-stu-id="85836-118">Requirements</span></span>

|<span data-ttu-id="85836-119">要求</span><span class="sxs-lookup"><span data-stu-id="85836-119">Requirement</span></span>| <span data-ttu-id="85836-120">值</span><span class="sxs-lookup"><span data-stu-id="85836-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="85836-121">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="85836-121">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="85836-122">1.0</span><span class="sxs-lookup"><span data-stu-id="85836-122">1.0</span></span>|
|[<span data-ttu-id="85836-123">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="85836-123">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="85836-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="85836-124">ReadItem</span></span>|
|[<span data-ttu-id="85836-125">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="85836-125">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="85836-126">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="85836-126">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="85836-127">示例</span><span class="sxs-lookup"><span data-stu-id="85836-127">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="85836-128">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="85836-128">emailAddress :String</span></span>

<span data-ttu-id="85836-129">获取用户的 SMTP 电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="85836-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="85836-130">类型：</span><span class="sxs-lookup"><span data-stu-id="85836-130">Type:</span></span>

*   <span data-ttu-id="85836-131">String</span><span class="sxs-lookup"><span data-stu-id="85836-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="85836-132">要求</span><span class="sxs-lookup"><span data-stu-id="85836-132">Requirements</span></span>

|<span data-ttu-id="85836-133">要求</span><span class="sxs-lookup"><span data-stu-id="85836-133">Requirement</span></span>| <span data-ttu-id="85836-134">值</span><span class="sxs-lookup"><span data-stu-id="85836-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="85836-135">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="85836-135">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="85836-136">1.0</span><span class="sxs-lookup"><span data-stu-id="85836-136">1.0</span></span>|
|[<span data-ttu-id="85836-137">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="85836-137">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="85836-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="85836-138">ReadItem</span></span>|
|[<span data-ttu-id="85836-139">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="85836-139">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="85836-140">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="85836-140">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="85836-141">示例</span><span class="sxs-lookup"><span data-stu-id="85836-141">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="85836-142">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="85836-142">timeZone :String</span></span>

<span data-ttu-id="85836-143">获取用户的默认时区。</span><span class="sxs-lookup"><span data-stu-id="85836-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="85836-144">类型：</span><span class="sxs-lookup"><span data-stu-id="85836-144">Type:</span></span>

*   <span data-ttu-id="85836-145">String</span><span class="sxs-lookup"><span data-stu-id="85836-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="85836-146">要求</span><span class="sxs-lookup"><span data-stu-id="85836-146">Requirements</span></span>

|<span data-ttu-id="85836-147">要求</span><span class="sxs-lookup"><span data-stu-id="85836-147">Requirement</span></span>| <span data-ttu-id="85836-148">值</span><span class="sxs-lookup"><span data-stu-id="85836-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="85836-149">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="85836-149">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="85836-150">1.0</span><span class="sxs-lookup"><span data-stu-id="85836-150">1.0</span></span>|
|[<span data-ttu-id="85836-151">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="85836-151">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="85836-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="85836-152">ReadItem</span></span>|
|[<span data-ttu-id="85836-153">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="85836-153">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="85836-154">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="85836-154">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="85836-155">示例</span><span class="sxs-lookup"><span data-stu-id="85836-155">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```