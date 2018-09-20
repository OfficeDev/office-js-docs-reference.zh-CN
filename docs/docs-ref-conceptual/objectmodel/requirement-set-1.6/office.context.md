
# <a name="context"></a><span data-ttu-id="435f4-101">context</span><span class="sxs-lookup"><span data-stu-id="435f4-101">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="435f4-102">[Office](Office.md).context</span><span class="sxs-lookup"><span data-stu-id="435f4-102">[Office](Office.md).context</span></span>

<span data-ttu-id="435f4-p101">Office.context 命名空间提供所有 Office 应用中的外接程序所使用的共享接口。此列表仅记录 Outlook 外接程序所使用的接口。有关 Office.context 命名空间的完整列表，请参阅[共享 API 中的 Office.context 引用](/javascript/api/office/office.context)。</span><span class="sxs-lookup"><span data-stu-id="435f4-p101">The Office.context namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Shared API](/javascript/api/office/office.context).</span></span>

##### <a name="requirements"></a><span data-ttu-id="435f4-105">要求</span><span class="sxs-lookup"><span data-stu-id="435f4-105">Requirements</span></span>

|<span data-ttu-id="435f4-106">要求</span><span class="sxs-lookup"><span data-stu-id="435f4-106">Requirement</span></span>| <span data-ttu-id="435f4-107">值</span><span class="sxs-lookup"><span data-stu-id="435f4-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="435f4-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="435f4-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="435f4-109">1.0</span><span class="sxs-lookup"><span data-stu-id="435f4-109">1.0</span></span>|
|[<span data-ttu-id="435f4-110">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="435f4-110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="435f4-111">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="435f4-111">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="435f4-112">成员和方法</span><span class="sxs-lookup"><span data-stu-id="435f4-112">Members and methods</span></span>

| <span data-ttu-id="435f4-113">成员</span><span class="sxs-lookup"><span data-stu-id="435f4-113">Member</span></span> | <span data-ttu-id="435f4-114">类型</span><span class="sxs-lookup"><span data-stu-id="435f4-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="435f4-115">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="435f4-115">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="435f4-116">成员</span><span class="sxs-lookup"><span data-stu-id="435f4-116">Member</span></span> |
| [<span data-ttu-id="435f4-117">officeTheme</span><span class="sxs-lookup"><span data-stu-id="435f4-117">officeTheme</span></span>](#officetheme-object) | <span data-ttu-id="435f4-118">成员</span><span class="sxs-lookup"><span data-stu-id="435f4-118">Member</span></span> |
| [<span data-ttu-id="435f4-119">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="435f4-119">roamingSettings</span></span>](#roamingsettings-roamingsettingsjavascriptapioutlook16officeroamingsettings) | <span data-ttu-id="435f4-120">成员</span><span class="sxs-lookup"><span data-stu-id="435f4-120">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="435f4-121">命名空间</span><span class="sxs-lookup"><span data-stu-id="435f4-121">Namespaces</span></span>

<span data-ttu-id="435f4-122">[邮箱](office.context.mailbox.md)： Microsoft Outlook 和 Microsoft Outlook on the web 提供对 Outlook 加载项对象模型的访问。</span><span class="sxs-lookup"><span data-stu-id="435f4-122">[mailbox](office.context.mailbox.md): Provides access to the Outlook add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

### <a name="members"></a><span data-ttu-id="435f4-123">成员</span><span class="sxs-lookup"><span data-stu-id="435f4-123">Members</span></span>

####  <a name="displaylanguage-string"></a><span data-ttu-id="435f4-124">displayLanguage :String</span><span class="sxs-lookup"><span data-stu-id="435f4-124">displayLanguage :String</span></span>

<span data-ttu-id="435f4-125">获取用户针对 Office 主机应用程序的 UI 指定的 RFC 1766 语言标记格式的区域设置（语言）。</span><span class="sxs-lookup"><span data-stu-id="435f4-125">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="435f4-126">`displayLanguage` 值反映在 Office 主机应用程序中通过“**文件 > 选项 > 语言**”指定的当前“**显示语言**”设置。</span><span class="sxs-lookup"><span data-stu-id="435f4-126">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="435f4-127">类型：</span><span class="sxs-lookup"><span data-stu-id="435f4-127">Type:</span></span>

*   <span data-ttu-id="435f4-128">String</span><span class="sxs-lookup"><span data-stu-id="435f4-128">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="435f4-129">要求</span><span class="sxs-lookup"><span data-stu-id="435f4-129">Requirements</span></span>

|<span data-ttu-id="435f4-130">要求</span><span class="sxs-lookup"><span data-stu-id="435f4-130">Requirement</span></span>| <span data-ttu-id="435f4-131">值</span><span class="sxs-lookup"><span data-stu-id="435f4-131">Value</span></span>|
|---|---|
|[<span data-ttu-id="435f4-132">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="435f4-132">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="435f4-133">1.0</span><span class="sxs-lookup"><span data-stu-id="435f4-133">1.0</span></span>|
|[<span data-ttu-id="435f4-134">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="435f4-134">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="435f4-135">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="435f4-135">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="435f4-136">示例</span><span class="sxs-lookup"><span data-stu-id="435f4-136">Example</span></span>

```js
function sayHelloWithDisplayLanguage() {
  var myDisplayLanguage = Office.context.displayLanguage;
  switch (myDisplayLanguage) {
    case 'en-US':
      write('Hello!');
      break;
    case 'en-NZ':
      write('G\'day mate!');
      break;
  }
}
// Function that writes to a div with id='message' on the page.
function write(message){
  document.getElementById('message').innerText += message;
}
```

####  <a name="officetheme-object"></a><span data-ttu-id="435f4-137">officeTheme :Object</span><span class="sxs-lookup"><span data-stu-id="435f4-137">officeTheme :Object</span></span>

<span data-ttu-id="435f4-138">提供了访问 Office 主题颜色的属性。</span><span class="sxs-lookup"><span data-stu-id="435f4-138">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="435f4-139">此成员不支持适用于 iOS 的 Outlook 或 Outlook for Android。</span><span class="sxs-lookup"><span data-stu-id="435f4-139">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="435f4-p102">通过使用 Office 主题颜色，你可以使外接程序的配色方案与用户（通过 **“文件”>“Office 帐户”>“Office 主题”UI**）选择的当前 Office 主题协调一致，这种做法适用于所有 Office 主机应用程序。使用 Office 主题颜色适用于邮件和任务窗格外接程序。</span><span class="sxs-lookup"><span data-stu-id="435f4-p102">Using Office theme colors let's you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office host applications. Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="435f4-142">类型：</span><span class="sxs-lookup"><span data-stu-id="435f4-142">Type:</span></span>

*   <span data-ttu-id="435f4-143">对象</span><span class="sxs-lookup"><span data-stu-id="435f4-143">Object</span></span>

##### <a name="properties"></a><span data-ttu-id="435f4-144">属性：</span><span class="sxs-lookup"><span data-stu-id="435f4-144">Properties:</span></span>

|<span data-ttu-id="435f4-145">名称</span><span class="sxs-lookup"><span data-stu-id="435f4-145">Name</span></span>| <span data-ttu-id="435f4-146">类型</span><span class="sxs-lookup"><span data-stu-id="435f4-146">Type</span></span>| <span data-ttu-id="435f4-147">说明</span><span class="sxs-lookup"><span data-stu-id="435f4-147">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="435f4-148">String</span><span class="sxs-lookup"><span data-stu-id="435f4-148">String</span></span>|<span data-ttu-id="435f4-149">获取十六进制三原色形式的 Office 主题正文背景色。</span><span class="sxs-lookup"><span data-stu-id="435f4-149">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="435f4-150">String</span><span class="sxs-lookup"><span data-stu-id="435f4-150">String</span></span>|<span data-ttu-id="435f4-151">获取十六进制三原色形式的 Office 主题正文前景色。</span><span class="sxs-lookup"><span data-stu-id="435f4-151">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="435f4-152">字符串</span><span class="sxs-lookup"><span data-stu-id="435f4-152">String</span></span>|<span data-ttu-id="435f4-153">获取十六进制三原色形式的 Office 主题控制背景色。</span><span class="sxs-lookup"><span data-stu-id="435f4-153">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="435f4-154">字符串</span><span class="sxs-lookup"><span data-stu-id="435f4-154">String</span></span>|<span data-ttu-id="435f4-155">获取十六进制三原色形式的 Office 主题正文控制颜色。</span><span class="sxs-lookup"><span data-stu-id="435f4-155">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="435f4-156">要求</span><span class="sxs-lookup"><span data-stu-id="435f4-156">Requirements</span></span>

|<span data-ttu-id="435f4-157">要求</span><span class="sxs-lookup"><span data-stu-id="435f4-157">Requirement</span></span>| <span data-ttu-id="435f4-158">值</span><span class="sxs-lookup"><span data-stu-id="435f4-158">Value</span></span>|
|---|---|
|[<span data-ttu-id="435f4-159">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="435f4-159">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="435f4-160">1.3</span><span class="sxs-lookup"><span data-stu-id="435f4-160">1.3</span></span>|
|[<span data-ttu-id="435f4-161">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="435f4-161">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="435f4-162">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="435f4-162">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="435f4-163">示例</span><span class="sxs-lookup"><span data-stu-id="435f4-163">Example</span></span>

```js
function applyOfficeTheme(){
  // Get office theme colors.
  var bodyBackgroundColor = Office.context.officeTheme.bodyBackgroundColor;
  var bodyForegroundColor = Office.context.officeTheme.bodyForegroundColor;
  var controlBackgroundColor = Office.context.officeTheme.controlBackgroundColor
  var controlForegroundColor = Office.context.officeTheme.controlForegroundColor;

  // Apply body background color to a CSS class.
  $('.body').css('background-color', bodyBackgroundColor);
}
```

####  <a name="roamingsettings-roamingsettingsjavascriptapioutlook16officeroamingsettings"></a><span data-ttu-id="435f4-164">roamingSettings :[RoamingSettings](/javascript/api/outlook_1_6/office.RoamingSettings)</span><span class="sxs-lookup"><span data-stu-id="435f4-164">roamingSettings :[RoamingSettings](/javascript/api/outlook_1_6/office.RoamingSettings)</span></span>

<span data-ttu-id="435f4-165">获取一个对象，它表示保存到用户邮箱的邮件外接程序的自定义设置或状态。</span><span class="sxs-lookup"><span data-stu-id="435f4-165">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="435f4-166">`RoamingSettings` 对象允许您存储和访问用户邮箱中存储的邮件外接程序的数据，以便从用于访问该邮箱的任何主机客户端应用程序中运行该外接程序时，该外接程序可以使用该数据。</span><span class="sxs-lookup"><span data-stu-id="435f4-166">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="435f4-167">类型:</span><span class="sxs-lookup"><span data-stu-id="435f4-167">Type:</span></span>

*   [<span data-ttu-id="435f4-168">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="435f4-168">RoamingSettings</span></span>](/javascript/api/outlook_1_6/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="435f4-169">要求</span><span class="sxs-lookup"><span data-stu-id="435f4-169">Requirements</span></span>

|<span data-ttu-id="435f4-170">要求</span><span class="sxs-lookup"><span data-stu-id="435f4-170">Requirement</span></span>| <span data-ttu-id="435f4-171">值</span><span class="sxs-lookup"><span data-stu-id="435f4-171">Value</span></span>|
|---|---|
|[<span data-ttu-id="435f4-172">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="435f4-172">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="435f4-173">1.0</span><span class="sxs-lookup"><span data-stu-id="435f4-173">1.0</span></span>|
|[<span data-ttu-id="435f4-174">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="435f4-174">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="435f4-175">受限</span><span class="sxs-lookup"><span data-stu-id="435f4-175">Restricted</span></span>|
|[<span data-ttu-id="435f4-176">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="435f4-176">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="435f4-177">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="435f4-177">Compose or read</span></span>|