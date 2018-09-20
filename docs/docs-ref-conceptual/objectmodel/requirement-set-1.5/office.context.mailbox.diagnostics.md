# <a name="diagnostics"></a><span data-ttu-id="f10d3-101">diagnostics</span><span class="sxs-lookup"><span data-stu-id="f10d3-101">diagnostics</span></span>

### <span data-ttu-id="f10d3-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). diagnostics</span><span class="sxs-lookup"><span data-stu-id="f10d3-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). diagnostics</span></span>

<span data-ttu-id="f10d3-104">将诊断信息提供给 Outlook 外接程序。</span><span class="sxs-lookup"><span data-stu-id="f10d3-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="f10d3-105">要求</span><span class="sxs-lookup"><span data-stu-id="f10d3-105">Requirements</span></span>

|<span data-ttu-id="f10d3-106">要求</span><span class="sxs-lookup"><span data-stu-id="f10d3-106">Requirement</span></span>| <span data-ttu-id="f10d3-107">值</span><span class="sxs-lookup"><span data-stu-id="f10d3-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="f10d3-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f10d3-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f10d3-109">1.0</span><span class="sxs-lookup"><span data-stu-id="f10d3-109">1.0</span></span>|
|[<span data-ttu-id="f10d3-110">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f10d3-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f10d3-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f10d3-111">ReadItem</span></span>|
|[<span data-ttu-id="f10d3-112">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f10d3-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f10d3-113">Compose 或 Read</span><span class="sxs-lookup"><span data-stu-id="f10d3-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="f10d3-114">成员和方法</span><span class="sxs-lookup"><span data-stu-id="f10d3-114">Members and methods</span></span>

| <span data-ttu-id="f10d3-115">成员</span><span class="sxs-lookup"><span data-stu-id="f10d3-115">Member</span></span> | <span data-ttu-id="f10d3-116">类型</span><span class="sxs-lookup"><span data-stu-id="f10d3-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="f10d3-117">hostName</span><span class="sxs-lookup"><span data-stu-id="f10d3-117">hostName</span></span>](#hostname-string) | <span data-ttu-id="f10d3-118">成员</span><span class="sxs-lookup"><span data-stu-id="f10d3-118">Member</span></span> |
| [<span data-ttu-id="f10d3-119">hostVersion</span><span class="sxs-lookup"><span data-stu-id="f10d3-119">hostVersion</span></span>](#hostversion-string) | <span data-ttu-id="f10d3-120">成员</span><span class="sxs-lookup"><span data-stu-id="f10d3-120">Member</span></span> |
| [<span data-ttu-id="f10d3-121">OWAView</span><span class="sxs-lookup"><span data-stu-id="f10d3-121">OWAView</span></span>](#owaview-string) | <span data-ttu-id="f10d3-122">成员</span><span class="sxs-lookup"><span data-stu-id="f10d3-122">Member</span></span> |

### <a name="members"></a><span data-ttu-id="f10d3-123">成员</span><span class="sxs-lookup"><span data-stu-id="f10d3-123">Members</span></span>

####  <a name="hostname-string"></a><span data-ttu-id="f10d3-124">hostName :String</span><span class="sxs-lookup"><span data-stu-id="f10d3-124">hostName :String</span></span>

<span data-ttu-id="f10d3-125">获取表示主机应用程序的名称的字符串。</span><span class="sxs-lookup"><span data-stu-id="f10d3-125">Gets a string that represents the name of the host application.</span></span>

<span data-ttu-id="f10d3-126">一个 string 类型的值可以是下列值之一： `Outlook`， `OutlookIOS`，或`OutlookWebApp`。</span><span class="sxs-lookup"><span data-stu-id="f10d3-126">A string that can be one of the following values: `Outlook`, `OutlookIOS`, or `OutlookWebApp`.</span></span>

##### <a name="type"></a><span data-ttu-id="f10d3-127">类型:</span><span class="sxs-lookup"><span data-stu-id="f10d3-127">Type:</span></span>

*   <span data-ttu-id="f10d3-128">String</span><span class="sxs-lookup"><span data-stu-id="f10d3-128">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f10d3-129">要求</span><span class="sxs-lookup"><span data-stu-id="f10d3-129">Requirements</span></span>

|<span data-ttu-id="f10d3-130">要求</span><span class="sxs-lookup"><span data-stu-id="f10d3-130">Requirement</span></span>| <span data-ttu-id="f10d3-131">值</span><span class="sxs-lookup"><span data-stu-id="f10d3-131">Value</span></span>|
|---|---|
|[<span data-ttu-id="f10d3-132">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f10d3-132">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f10d3-133">1.0</span><span class="sxs-lookup"><span data-stu-id="f10d3-133">1.0</span></span>|
|[<span data-ttu-id="f10d3-134">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f10d3-134">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f10d3-135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f10d3-135">ReadItem</span></span>|
|[<span data-ttu-id="f10d3-136">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f10d3-136">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f10d3-137">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="f10d3-137">Compose or read</span></span>|

####  <a name="hostversion-string"></a><span data-ttu-id="f10d3-138">hostVersion :String</span><span class="sxs-lookup"><span data-stu-id="f10d3-138">hostVersion :String</span></span>

<span data-ttu-id="f10d3-139">获取表示主机应用程序或 Exchange Server 的版本的字符串。</span><span class="sxs-lookup"><span data-stu-id="f10d3-139">Gets a string that represents the version of either the host application or the Exchange Server.</span></span>

<span data-ttu-id="f10d3-p102">如果邮件外接程序正在 Outlook 桌面客户端或 Outlook for iOS 上运行，则 `hostVersion` 属性返回主机应用程序版本 Outlook。在 Outlook Web App 中，属性返回 Exchange Server 的版本。其中的一个示例是字符串 `15.0.468.0`。</span><span class="sxs-lookup"><span data-stu-id="f10d3-p102">If the mail add-in is running on the Outlook desktop client or Outlook for iOS, the `hostVersion` property returns the version of the host application, Outlook. In Outlook Web App, the property returns the version of the Exchange Server. An example is the string `15.0.468.0`.</span></span>

##### <a name="type"></a><span data-ttu-id="f10d3-143">类型:</span><span class="sxs-lookup"><span data-stu-id="f10d3-143">Type:</span></span>

*   <span data-ttu-id="f10d3-144">String</span><span class="sxs-lookup"><span data-stu-id="f10d3-144">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f10d3-145">要求</span><span class="sxs-lookup"><span data-stu-id="f10d3-145">Requirements</span></span>

|<span data-ttu-id="f10d3-146">要求</span><span class="sxs-lookup"><span data-stu-id="f10d3-146">Requirement</span></span>| <span data-ttu-id="f10d3-147">值</span><span class="sxs-lookup"><span data-stu-id="f10d3-147">Value</span></span>|
|---|---|
|[<span data-ttu-id="f10d3-148">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f10d3-148">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f10d3-149">1.0</span><span class="sxs-lookup"><span data-stu-id="f10d3-149">1.0</span></span>|
|[<span data-ttu-id="f10d3-150">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f10d3-150">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f10d3-151">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f10d3-151">ReadItem</span></span>|
|[<span data-ttu-id="f10d3-152">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f10d3-152">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f10d3-153">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="f10d3-153">Compose or read</span></span>|

####  <a name="owaview-string"></a><span data-ttu-id="f10d3-154">OWAView :String</span><span class="sxs-lookup"><span data-stu-id="f10d3-154">OWAView :String</span></span>

<span data-ttu-id="f10d3-155">获取表示 Outlook Web App 的当前视图的字符串。</span><span class="sxs-lookup"><span data-stu-id="f10d3-155">Gets a string that represents the current view of Outlook Web App.</span></span>

<span data-ttu-id="f10d3-156">返回的字符串可以是下列值之一：`OneColumn`、`TwoColumns` 或 `ThreeColumns`。</span><span class="sxs-lookup"><span data-stu-id="f10d3-156">The returned string can be one of the following values: `OneColumn`, `TwoColumns`, or `ThreeColumns`.</span></span>

<span data-ttu-id="f10d3-157">如果主机应用程序不是 Outlook Web App，则访问此属性将导致返回 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="f10d3-157">If the host application is not Outlook Web App, then accessing this property results in `undefined`.</span></span>

<span data-ttu-id="f10d3-158">Outlook Web App 具有三种视图，这些视图分别与屏幕和窗口的宽度以及可以显示的列数相对应：</span><span class="sxs-lookup"><span data-stu-id="f10d3-158">Outlook Web App has three views that correspond to the width of the screen and the window, and the number of columns that can be displayed:</span></span>

*   <span data-ttu-id="f10d3-p103">`OneColumn` 在屏幕较窄时显示。Outlook Web App 在智能手机的整个屏幕上使用此单列布局。</span><span class="sxs-lookup"><span data-stu-id="f10d3-p103">`OneColumn`, which is displayed when the screen is narrow. Outlook Web App uses this single-column layout on the entire screen of a smartphone.</span></span>
*   <span data-ttu-id="f10d3-p104">`TwoColumns` 在屏幕较宽时显示。Outlook Web App 在大多数平板电脑上使用此视图。</span><span class="sxs-lookup"><span data-stu-id="f10d3-p104">`TwoColumns`, which is displayed when the screen is wider. Outlook Web App uses this view on most tablets.</span></span>
*   <span data-ttu-id="f10d3-p105">`ThreeColumns` 在屏幕为宽屏时显示。例如，Outlook Web App 在台式机的全屏窗口中使用此视图。</span><span class="sxs-lookup"><span data-stu-id="f10d3-p105">`ThreeColumns`, which is displayed when the screen is wide. For example, Outlook Web App uses this view in a full screen window on a desktop computer.</span></span>

##### <a name="type"></a><span data-ttu-id="f10d3-165">类型：</span><span class="sxs-lookup"><span data-stu-id="f10d3-165">Type:</span></span>

*   <span data-ttu-id="f10d3-166">String</span><span class="sxs-lookup"><span data-stu-id="f10d3-166">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f10d3-167">要求</span><span class="sxs-lookup"><span data-stu-id="f10d3-167">Requirements</span></span>

|<span data-ttu-id="f10d3-168">要求</span><span class="sxs-lookup"><span data-stu-id="f10d3-168">Requirement</span></span>| <span data-ttu-id="f10d3-169">值</span><span class="sxs-lookup"><span data-stu-id="f10d3-169">Value</span></span>|
|---|---|
|[<span data-ttu-id="f10d3-170">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f10d3-170">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f10d3-171">1.0</span><span class="sxs-lookup"><span data-stu-id="f10d3-171">1.0</span></span>|
|[<span data-ttu-id="f10d3-172">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f10d3-172">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f10d3-173">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f10d3-173">ReadItem</span></span>|
|[<span data-ttu-id="f10d3-174">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f10d3-174">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f10d3-175">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="f10d3-175">Compose or read</span></span>|