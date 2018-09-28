
# <a name="mailbox"></a><span data-ttu-id="43f5c-101">mailbox</span><span class="sxs-lookup"><span data-stu-id="43f5c-101">mailbox</span></span>

### <span data-ttu-id="43f5c-p101">[Office](Office.md)[.context](Office.context.md). mailbox</span><span class="sxs-lookup"><span data-stu-id="43f5c-p101">[Office](Office.md)[.context](Office.context.md). mailbox</span></span>

<span data-ttu-id="43f5c-104">Microsoft Outlook 和 Microsoft Outlook on the web 提供对 Outlook 加载项对象模型的访问。</span><span class="sxs-lookup"><span data-stu-id="43f5c-104">Provides access to the Outlook add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

##### <a name="requirements"></a><span data-ttu-id="43f5c-105">要求</span><span class="sxs-lookup"><span data-stu-id="43f5c-105">Requirements</span></span>

|<span data-ttu-id="43f5c-106">要求</span><span class="sxs-lookup"><span data-stu-id="43f5c-106">Requirement</span></span>| <span data-ttu-id="43f5c-107">值</span><span class="sxs-lookup"><span data-stu-id="43f5c-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="43f5c-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="43f5c-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43f5c-109">1.0</span><span class="sxs-lookup"><span data-stu-id="43f5c-109">1.0</span></span>|
|[<span data-ttu-id="43f5c-110">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="43f5c-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43f5c-111">受限</span><span class="sxs-lookup"><span data-stu-id="43f5c-111">Restricted</span></span>|
|[<span data-ttu-id="43f5c-112">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="43f5c-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="43f5c-113">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="43f5c-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="43f5c-114">成员和方法</span><span class="sxs-lookup"><span data-stu-id="43f5c-114">Members and methods</span></span>

| <span data-ttu-id="43f5c-115">成员</span><span class="sxs-lookup"><span data-stu-id="43f5c-115">Member</span></span> | <span data-ttu-id="43f5c-116">类型</span><span class="sxs-lookup"><span data-stu-id="43f5c-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="43f5c-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="43f5c-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="43f5c-118">成员</span><span class="sxs-lookup"><span data-stu-id="43f5c-118">Member</span></span> |
| [<span data-ttu-id="43f5c-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="43f5c-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="43f5c-120">成员</span><span class="sxs-lookup"><span data-stu-id="43f5c-120">Member</span></span> |
| [<span data-ttu-id="43f5c-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="43f5c-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="43f5c-122">方法</span><span class="sxs-lookup"><span data-stu-id="43f5c-122">Method</span></span> |
| [<span data-ttu-id="43f5c-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="43f5c-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="43f5c-124">方法</span><span class="sxs-lookup"><span data-stu-id="43f5c-124">Method</span></span> |
| [<span data-ttu-id="43f5c-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="43f5c-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime) | <span data-ttu-id="43f5c-126">方法</span><span class="sxs-lookup"><span data-stu-id="43f5c-126">Method</span></span> |
| [<span data-ttu-id="43f5c-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="43f5c-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="43f5c-128">方法</span><span class="sxs-lookup"><span data-stu-id="43f5c-128">Method</span></span> |
| [<span data-ttu-id="43f5c-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="43f5c-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="43f5c-130">方法</span><span class="sxs-lookup"><span data-stu-id="43f5c-130">Method</span></span> |
| [<span data-ttu-id="43f5c-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="43f5c-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="43f5c-132">方法</span><span class="sxs-lookup"><span data-stu-id="43f5c-132">Method</span></span> |
| [<span data-ttu-id="43f5c-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="43f5c-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="43f5c-134">方法</span><span class="sxs-lookup"><span data-stu-id="43f5c-134">Method</span></span> |
| [<span data-ttu-id="43f5c-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="43f5c-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="43f5c-136">方法</span><span class="sxs-lookup"><span data-stu-id="43f5c-136">Method</span></span> |
| [<span data-ttu-id="43f5c-137">displayNewMessageForm</span><span class="sxs-lookup"><span data-stu-id="43f5c-137">displayNewMessageForm</span></span>](#displaynewmessageformparameters) | <span data-ttu-id="43f5c-138">方法</span><span class="sxs-lookup"><span data-stu-id="43f5c-138">Method</span></span> |
| [<span data-ttu-id="43f5c-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="43f5c-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="43f5c-140">方法</span><span class="sxs-lookup"><span data-stu-id="43f5c-140">Method</span></span> |
| [<span data-ttu-id="43f5c-141">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="43f5c-141">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="43f5c-142">方法</span><span class="sxs-lookup"><span data-stu-id="43f5c-142">Method</span></span> |
| [<span data-ttu-id="43f5c-143">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="43f5c-143">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="43f5c-144">方法</span><span class="sxs-lookup"><span data-stu-id="43f5c-144">Method</span></span> |
| [<span data-ttu-id="43f5c-145">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="43f5c-145">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="43f5c-146">方法</span><span class="sxs-lookup"><span data-stu-id="43f5c-146">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="43f5c-147">命名空间</span><span class="sxs-lookup"><span data-stu-id="43f5c-147">Namespaces</span></span>

<span data-ttu-id="43f5c-148">[diagnostics](Office.context.mailbox.diagnostics.md)：将诊断信息提供给 Outlook 外接程序。</span><span class="sxs-lookup"><span data-stu-id="43f5c-148">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="43f5c-149">[item](Office.context.mailbox.item.md)：提供用于访问 Outlook 外接程序中的邮件或约会的方法和属性。</span><span class="sxs-lookup"><span data-stu-id="43f5c-149">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="43f5c-150">[userProfile](Office.context.mailbox.userProfile.md)：提供有关 Outlook 外接程序中的用户的信息。</span><span class="sxs-lookup"><span data-stu-id="43f5c-150">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="43f5c-151">成员</span><span class="sxs-lookup"><span data-stu-id="43f5c-151">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="43f5c-152">ewsUrl :String</span><span class="sxs-lookup"><span data-stu-id="43f5c-152">ewsUrl :String</span></span>

<span data-ttu-id="43f5c-p102">获取此电子邮件帐户的 Exchange Web 服务 (EWS) 终结点的 URL。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="43f5c-p102">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="43f5c-155">此成员不支持适用于 iOS 的 Outlook 或 Outlook for Android。</span><span class="sxs-lookup"><span data-stu-id="43f5c-155">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="43f5c-p103">远程服务可使用 `ewsUrl` 值对用户邮箱进行 EWS 调用。例如，可以创建远程服务来 [获取选定项目中的附件](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item)。</span><span class="sxs-lookup"><span data-stu-id="43f5c-p103">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="43f5c-158">应用必须在其清单中指定拥有 **ReadItem** 权限，才能调用阅读模式中的 `ewsUrl` 成员。</span><span class="sxs-lookup"><span data-stu-id="43f5c-158">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="43f5c-p104">在撰写模式中，必须调用 [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) 方法，才能使用 `ewsUrl` 成员。应用必须具有调用 `saveAsync` 方法的 **ReadWriteItem** 权限。</span><span class="sxs-lookup"><span data-stu-id="43f5c-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="43f5c-161">类型:</span><span class="sxs-lookup"><span data-stu-id="43f5c-161">Type:</span></span>

*   <span data-ttu-id="43f5c-162">String</span><span class="sxs-lookup"><span data-stu-id="43f5c-162">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="43f5c-163">要求</span><span class="sxs-lookup"><span data-stu-id="43f5c-163">Requirements</span></span>

|<span data-ttu-id="43f5c-164">要求</span><span class="sxs-lookup"><span data-stu-id="43f5c-164">Requirement</span></span>| <span data-ttu-id="43f5c-165">值</span><span class="sxs-lookup"><span data-stu-id="43f5c-165">Value</span></span>|
|---|---|
|[<span data-ttu-id="43f5c-166">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="43f5c-166">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43f5c-167">1.0</span><span class="sxs-lookup"><span data-stu-id="43f5c-167">1.0</span></span>|
|[<span data-ttu-id="43f5c-168">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="43f5c-168">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43f5c-169">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43f5c-169">ReadItem</span></span>|
|[<span data-ttu-id="43f5c-170">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="43f5c-170">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="43f5c-171">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="43f5c-171">Compose or read</span></span>|

#### <a name="resturl-string"></a><span data-ttu-id="43f5c-172">restUrl :String</span><span class="sxs-lookup"><span data-stu-id="43f5c-172">restUrl :String</span></span>

<span data-ttu-id="43f5c-173">获取此电子邮件帐户的 REST 终结点的 URL。</span><span class="sxs-lookup"><span data-stu-id="43f5c-173">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="43f5c-174">`restUrl` 值可用于对用户邮箱进行 [REST API](https://docs.microsoft.com/outlook/rest/) 调用。</span><span class="sxs-lookup"><span data-stu-id="43f5c-174">The `restUrl` value can be used to make [REST API](https://docs.microsoft.com/outlook/rest/) calls to the user's mailbox.</span></span>

<span data-ttu-id="43f5c-175">应用必须在其清单中指定拥有 **ReadItem** 权限，才能调用阅读模式中的 `restUrl` 成员。</span><span class="sxs-lookup"><span data-stu-id="43f5c-175">Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.</span></span>

<span data-ttu-id="43f5c-p105">在撰写模式中，必须调用 [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) 方法，才能使用 `restUrl` 成员。应用必须具有调用 `saveAsync` 方法的 **ReadWriteItem** 权限。</span><span class="sxs-lookup"><span data-stu-id="43f5c-p105">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="43f5c-178">类型:</span><span class="sxs-lookup"><span data-stu-id="43f5c-178">Type:</span></span>

*   <span data-ttu-id="43f5c-179">String</span><span class="sxs-lookup"><span data-stu-id="43f5c-179">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="43f5c-180">要求</span><span class="sxs-lookup"><span data-stu-id="43f5c-180">Requirements</span></span>

|<span data-ttu-id="43f5c-181">要求</span><span class="sxs-lookup"><span data-stu-id="43f5c-181">Requirement</span></span>| <span data-ttu-id="43f5c-182">值</span><span class="sxs-lookup"><span data-stu-id="43f5c-182">Value</span></span>|
|---|---|
|[<span data-ttu-id="43f5c-183">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="43f5c-183">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43f5c-184">1.5</span><span class="sxs-lookup"><span data-stu-id="43f5c-184">1.5</span></span> |
|[<span data-ttu-id="43f5c-185">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="43f5c-185">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43f5c-186">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43f5c-186">ReadItem</span></span>|
|[<span data-ttu-id="43f5c-187">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="43f5c-187">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="43f5c-188">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="43f5c-188">Compose or read</span></span>|

### <a name="methods"></a><span data-ttu-id="43f5c-189">方法</span><span class="sxs-lookup"><span data-stu-id="43f5c-189">Methods</span></span>

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="43f5c-190">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="43f5c-190">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="43f5c-191">添加支持事件的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="43f5c-191">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="43f5c-p106">目前，唯一受支持的事件类型是 `Office.EventType.ItemChanged`，用户选择一个新项目时将调用该类型。此事件由实现可固定任务窗格的外接程序使用，并允许外接程序根据当前选定的项目刷新任务窗格 UI。</span><span class="sxs-lookup"><span data-stu-id="43f5c-p106">Currently the only supported event type is `Office.EventType.ItemChanged`, which is invoked when the user selects a new item. This event is used by add-ins that implement a pinnable taskpane, and allows the add-in to refresh the taskpane UI based on the currently selected item.</span></span>

##### <a name="parameters"></a><span data-ttu-id="43f5c-194">参数：</span><span class="sxs-lookup"><span data-stu-id="43f5c-194">Parameters:</span></span>

| <span data-ttu-id="43f5c-195">名称</span><span class="sxs-lookup"><span data-stu-id="43f5c-195">Name</span></span> | <span data-ttu-id="43f5c-196">类型</span><span class="sxs-lookup"><span data-stu-id="43f5c-196">Type</span></span> | <span data-ttu-id="43f5c-197">属性</span><span class="sxs-lookup"><span data-stu-id="43f5c-197">Attributes</span></span> | <span data-ttu-id="43f5c-198">说明</span><span class="sxs-lookup"><span data-stu-id="43f5c-198">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="43f5c-199">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="43f5c-199">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="43f5c-200">应调用处理程序的事件。</span><span class="sxs-lookup"><span data-stu-id="43f5c-200">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="43f5c-201">函数</span><span class="sxs-lookup"><span data-stu-id="43f5c-201">Function</span></span> || <span data-ttu-id="43f5c-p107">用于处理事件的函数。此函数必须接受一个参数，即对象文本。参数上的 `type` 属性将匹配传递给 `addHandlerAsync` 的 `eventType` 参数。</span><span class="sxs-lookup"><span data-stu-id="43f5c-p107">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="43f5c-205">Object</span><span class="sxs-lookup"><span data-stu-id="43f5c-205">Object</span></span> | <span data-ttu-id="43f5c-206">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="43f5c-206">&lt;optional&gt;</span></span> | <span data-ttu-id="43f5c-207">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="43f5c-207">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="43f5c-208">对象</span><span class="sxs-lookup"><span data-stu-id="43f5c-208">Object</span></span> | <span data-ttu-id="43f5c-209">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="43f5c-209">&lt;optional&gt;</span></span> | <span data-ttu-id="43f5c-210">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="43f5c-210">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="43f5c-211">函数</span><span class="sxs-lookup"><span data-stu-id="43f5c-211">function</span></span>| <span data-ttu-id="43f5c-212">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="43f5c-212">&lt;optional&gt;</span></span>|<span data-ttu-id="43f5c-213">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="43f5c-213">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="43f5c-214">Requirements</span><span class="sxs-lookup"><span data-stu-id="43f5c-214">Requirements</span></span>

|<span data-ttu-id="43f5c-215">要求</span><span class="sxs-lookup"><span data-stu-id="43f5c-215">Requirement</span></span>| <span data-ttu-id="43f5c-216">值</span><span class="sxs-lookup"><span data-stu-id="43f5c-216">Value</span></span>|
|---|---|
|[<span data-ttu-id="43f5c-217">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="43f5c-217">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43f5c-218">1.5</span><span class="sxs-lookup"><span data-stu-id="43f5c-218">1.5</span></span> |
|[<span data-ttu-id="43f5c-219">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="43f5c-219">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43f5c-220">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43f5c-220">ReadItem</span></span> |
|[<span data-ttu-id="43f5c-221">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="43f5c-221">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="43f5c-222">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="43f5c-222">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="43f5c-223">示例</span><span class="sxs-lookup"><span data-stu-id="43f5c-223">Example</span></span>

```
Office.initialize = function (reason) {
  $(document).ready(function () {
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, loadNewItem, function (result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        // Handle error
      }
    });
  });
};

function loadNewItem(eventArgs) {
  // Load the properties of the newly selected item
  loadProps(Office.context.mailbox.item);
};
```

####  <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="43f5c-224">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="43f5c-224">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="43f5c-225">将项目 ID 格式化（从 REST 转换为 EWS 格式）。</span><span class="sxs-lookup"><span data-stu-id="43f5c-225">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="43f5c-226">在适用于 iOS 的 Outlook 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="43f5c-226">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="43f5c-p108">通过 REST API 检索的项 ID（如 [Outlook 邮件 API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) 或 [Microsoft Graph](http://graph.microsoft.io/)）使用与 Exchange Web 服务 (EWS) 所使用格式不同的格式。`convertToEwsId` 方法将 REST 格式化的 ID 转换为正确的 EWS 格式。</span><span class="sxs-lookup"><span data-stu-id="43f5c-p108">Item IDs retrieved via a REST API (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](http://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="43f5c-229">参数：</span><span class="sxs-lookup"><span data-stu-id="43f5c-229">Parameters:</span></span>

|<span data-ttu-id="43f5c-230">名称</span><span class="sxs-lookup"><span data-stu-id="43f5c-230">Name</span></span>| <span data-ttu-id="43f5c-231">类型</span><span class="sxs-lookup"><span data-stu-id="43f5c-231">Type</span></span>| <span data-ttu-id="43f5c-232">说明</span><span class="sxs-lookup"><span data-stu-id="43f5c-232">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="43f5c-233">字符串</span><span class="sxs-lookup"><span data-stu-id="43f5c-233">String</span></span>|<span data-ttu-id="43f5c-234">Outlook REST API 的格式化的项目 ID。</span><span class="sxs-lookup"><span data-stu-id="43f5c-234">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="43f5c-235">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="43f5c-235">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.restversion)|<span data-ttu-id="43f5c-236">指示用于检索项目 ID 的 Outlook REST API 的版本。</span><span class="sxs-lookup"><span data-stu-id="43f5c-236">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="43f5c-237">要求</span><span class="sxs-lookup"><span data-stu-id="43f5c-237">Requirements</span></span>

|<span data-ttu-id="43f5c-238">要求</span><span class="sxs-lookup"><span data-stu-id="43f5c-238">Requirement</span></span>| <span data-ttu-id="43f5c-239">值</span><span class="sxs-lookup"><span data-stu-id="43f5c-239">Value</span></span>|
|---|---|
|[<span data-ttu-id="43f5c-240">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="43f5c-240">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43f5c-241">1.3</span><span class="sxs-lookup"><span data-stu-id="43f5c-241">1.3</span></span>|
|[<span data-ttu-id="43f5c-242">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="43f5c-242">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43f5c-243">受限</span><span class="sxs-lookup"><span data-stu-id="43f5c-243">Restricted</span></span>|
|[<span data-ttu-id="43f5c-244">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="43f5c-244">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="43f5c-245">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="43f5c-245">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="43f5c-246">返回：</span><span class="sxs-lookup"><span data-stu-id="43f5c-246">Returns:</span></span>

<span data-ttu-id="43f5c-247">类型：字符串</span><span class="sxs-lookup"><span data-stu-id="43f5c-247">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="43f5c-248">示例</span><span class="sxs-lookup"><span data-stu-id="43f5c-248">Example</span></span>

```
// Get an item's ID from a REST API
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the
// Outlook Mail API
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime"></a><span data-ttu-id="43f5c-249">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_6/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="43f5c-249">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_6/office.LocalClientTime)}</span></span>

<span data-ttu-id="43f5c-250">获取包含以本地客户端时间表示的时间信息的字典。</span><span class="sxs-lookup"><span data-stu-id="43f5c-250">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="43f5c-p109">Outlook 或 Outlook Web App 邮件应用程序的日期和时间可以使用不同的时区。Outlook 使用客户端计算机时区；Outlook Web App 使用 Exchange 管理中心 (EAC) 中设置的时区。应对日期和时间值进行处理，以便用户界面上显示的值始终与用户预期的时区一致。</span><span class="sxs-lookup"><span data-stu-id="43f5c-p109">The dates and times used by a mail app for Outlook or Outlook Web App can use different time zones. Outlook uses the client computer time zone; Outlook Web App uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="43f5c-p110">如果邮件应用程序在 Outlook 中运行，`convertToLocalClientTime` 方法将返回一个值设置为客户端计算机时区的字典对象。如果邮件应用程序在 Outlook Web App 中运行，`convertToLocalClientTime` 方法将返回值设置为 EAC 中指定的时区的字典对象。</span><span class="sxs-lookup"><span data-stu-id="43f5c-p110">If the mail app is running in Outlook, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook Web App, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="43f5c-256">参数：</span><span class="sxs-lookup"><span data-stu-id="43f5c-256">Parameters:</span></span>

|<span data-ttu-id="43f5c-257">名称</span><span class="sxs-lookup"><span data-stu-id="43f5c-257">Name</span></span>| <span data-ttu-id="43f5c-258">类型</span><span class="sxs-lookup"><span data-stu-id="43f5c-258">Type</span></span>| <span data-ttu-id="43f5c-259">说明</span><span class="sxs-lookup"><span data-stu-id="43f5c-259">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="43f5c-260">日期</span><span class="sxs-lookup"><span data-stu-id="43f5c-260">Date</span></span>|<span data-ttu-id="43f5c-261">一个 Date 对象</span><span class="sxs-lookup"><span data-stu-id="43f5c-261">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="43f5c-262">要求</span><span class="sxs-lookup"><span data-stu-id="43f5c-262">Requirements</span></span>

|<span data-ttu-id="43f5c-263">要求</span><span class="sxs-lookup"><span data-stu-id="43f5c-263">Requirement</span></span>| <span data-ttu-id="43f5c-264">值</span><span class="sxs-lookup"><span data-stu-id="43f5c-264">Value</span></span>|
|---|---|
|[<span data-ttu-id="43f5c-265">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="43f5c-265">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43f5c-266">1.0</span><span class="sxs-lookup"><span data-stu-id="43f5c-266">1.0</span></span>|
|[<span data-ttu-id="43f5c-267">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="43f5c-267">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43f5c-268">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43f5c-268">ReadItem</span></span>|
|[<span data-ttu-id="43f5c-269">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="43f5c-269">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="43f5c-270">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="43f5c-270">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="43f5c-271">返回：</span><span class="sxs-lookup"><span data-stu-id="43f5c-271">Returns:</span></span>

<span data-ttu-id="43f5c-272">类型：[LocalClientTime](/javascript/api/outlook_1_6/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="43f5c-272">Type: [LocalClientTime](/javascript/api/outlook_1_6/office.LocalClientTime)</span></span>

####  <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="43f5c-273">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="43f5c-273">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="43f5c-274">将项目 ID 格式化（从 EWS 转换为 REST 格式）。</span><span class="sxs-lookup"><span data-stu-id="43f5c-274">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="43f5c-275">在适用于 iOS 的 Outlook 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="43f5c-275">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="43f5c-p111">与 REST API 所使用的格式比较，通过 EWS 或通过 `itemId` 属性检索的项目 ID 使用不同的格式（例如 [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) 或 [Microsoft Graph](http://graph.microsoft.io/)）。`convertToRestId` 方法将 EWS 格式化的 ID 转换为正确的 REST 格式。</span><span class="sxs-lookup"><span data-stu-id="43f5c-p111">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](http://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="43f5c-278">参数：</span><span class="sxs-lookup"><span data-stu-id="43f5c-278">Parameters:</span></span>

|<span data-ttu-id="43f5c-279">名称</span><span class="sxs-lookup"><span data-stu-id="43f5c-279">Name</span></span>| <span data-ttu-id="43f5c-280">类型</span><span class="sxs-lookup"><span data-stu-id="43f5c-280">Type</span></span>| <span data-ttu-id="43f5c-281">说明</span><span class="sxs-lookup"><span data-stu-id="43f5c-281">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="43f5c-282">字符串</span><span class="sxs-lookup"><span data-stu-id="43f5c-282">String</span></span>|<span data-ttu-id="43f5c-283">适用于 Exchange Web 服务 (EWS) 的项目 ID 格式化。</span><span class="sxs-lookup"><span data-stu-id="43f5c-283">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="43f5c-284">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="43f5c-284">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.restversion)|<span data-ttu-id="43f5c-285">值指示转换的 ID 所使用的 Outlook REST API 的版本。</span><span class="sxs-lookup"><span data-stu-id="43f5c-285">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="43f5c-286">要求</span><span class="sxs-lookup"><span data-stu-id="43f5c-286">Requirements</span></span>

|<span data-ttu-id="43f5c-287">要求</span><span class="sxs-lookup"><span data-stu-id="43f5c-287">Requirement</span></span>| <span data-ttu-id="43f5c-288">值</span><span class="sxs-lookup"><span data-stu-id="43f5c-288">Value</span></span>|
|---|---|
|[<span data-ttu-id="43f5c-289">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="43f5c-289">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43f5c-290">1.3</span><span class="sxs-lookup"><span data-stu-id="43f5c-290">1.3</span></span>|
|[<span data-ttu-id="43f5c-291">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="43f5c-291">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43f5c-292">受限</span><span class="sxs-lookup"><span data-stu-id="43f5c-292">Restricted</span></span>|
|[<span data-ttu-id="43f5c-293">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="43f5c-293">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="43f5c-294">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="43f5c-294">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="43f5c-295">返回：</span><span class="sxs-lookup"><span data-stu-id="43f5c-295">Returns:</span></span>

<span data-ttu-id="43f5c-296">类型：字符串</span><span class="sxs-lookup"><span data-stu-id="43f5c-296">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="43f5c-297">示例</span><span class="sxs-lookup"><span data-stu-id="43f5c-297">Example</span></span>

```
// Get the currently selected item's ID
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the
// Outlook Mail API
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="43f5c-298">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="43f5c-298">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="43f5c-299">从包含时间信息的字典中获取 Date 对象。</span><span class="sxs-lookup"><span data-stu-id="43f5c-299">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="43f5c-300">`convertToUtcClientTime` 方法将包含本地日期和时间的字典转换为包含与本地日期和时间对应的正确值的 Date 对象。</span><span class="sxs-lookup"><span data-stu-id="43f5c-300">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="43f5c-301">参数：</span><span class="sxs-lookup"><span data-stu-id="43f5c-301">Parameters:</span></span>

|<span data-ttu-id="43f5c-302">名称</span><span class="sxs-lookup"><span data-stu-id="43f5c-302">Name</span></span>| <span data-ttu-id="43f5c-303">类型</span><span class="sxs-lookup"><span data-stu-id="43f5c-303">Type</span></span>| <span data-ttu-id="43f5c-304">说明</span><span class="sxs-lookup"><span data-stu-id="43f5c-304">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="43f5c-305">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="43f5c-305">LocalClientTime</span></span>](/javascript/api/outlook_1_6/office.LocalClientTime)|<span data-ttu-id="43f5c-306">要转换的本地时间值。</span><span class="sxs-lookup"><span data-stu-id="43f5c-306">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="43f5c-307">要求</span><span class="sxs-lookup"><span data-stu-id="43f5c-307">Requirements</span></span>

|<span data-ttu-id="43f5c-308">要求</span><span class="sxs-lookup"><span data-stu-id="43f5c-308">Requirement</span></span>| <span data-ttu-id="43f5c-309">值</span><span class="sxs-lookup"><span data-stu-id="43f5c-309">Value</span></span>|
|---|---|
|[<span data-ttu-id="43f5c-310">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="43f5c-310">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43f5c-311">1.0</span><span class="sxs-lookup"><span data-stu-id="43f5c-311">1.0</span></span>|
|[<span data-ttu-id="43f5c-312">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="43f5c-312">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43f5c-313">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43f5c-313">ReadItem</span></span>|
|[<span data-ttu-id="43f5c-314">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="43f5c-314">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="43f5c-315">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="43f5c-315">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="43f5c-316">返回：</span><span class="sxs-lookup"><span data-stu-id="43f5c-316">Returns:</span></span>

<span data-ttu-id="43f5c-317">包含以 UTC 表示的时间的 Date 对象。</span><span class="sxs-lookup"><span data-stu-id="43f5c-317">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="43f5c-318">

<dt>
类型</dt>


</span><span class="sxs-lookup"><span data-stu-id="43f5c-318">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="43f5c-319">日期</span><span class="sxs-lookup"><span data-stu-id="43f5c-319">Date</span></span></dd>

</dl>

####  <a name="displayappointmentformitemid"></a><span data-ttu-id="43f5c-320">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="43f5c-320">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="43f5c-321">显示现有日历约会。</span><span class="sxs-lookup"><span data-stu-id="43f5c-321">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="43f5c-322">在适用于 iOS 的 Outlook 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="43f5c-322">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="43f5c-323">`displayAppointmentForm` 方法将打开桌面新窗口中或移动设备对话框中的现有日历约会。</span><span class="sxs-lookup"><span data-stu-id="43f5c-323">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="43f5c-p112">在 Outlook for Mac 中，您可以使用此方法来显示不属于定期系列的单个约会，或显示定期系列的主约会，但无法显示该系列的实例。这是因为在 Outlook for Mac 中，无法访问定期系列实例的属性（包括项目 ID）。</span><span class="sxs-lookup"><span data-stu-id="43f5c-p112">In Outlook for Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook for Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="43f5c-326">在 Outlook Web App 中，此方法仅在窗体正文小于或等于 32KB 字符数时，才会打开指定的窗体。</span><span class="sxs-lookup"><span data-stu-id="43f5c-326">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="43f5c-327">如果指定的项标识符没有识别现有约会，将在客户端计算机或设备上打开一个空白窗格，并且不会返回错误消息。</span><span class="sxs-lookup"><span data-stu-id="43f5c-327">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="43f5c-328">参数：</span><span class="sxs-lookup"><span data-stu-id="43f5c-328">Parameters:</span></span>

|<span data-ttu-id="43f5c-329">名称</span><span class="sxs-lookup"><span data-stu-id="43f5c-329">Name</span></span>| <span data-ttu-id="43f5c-330">类型</span><span class="sxs-lookup"><span data-stu-id="43f5c-330">Type</span></span>| <span data-ttu-id="43f5c-331">说明</span><span class="sxs-lookup"><span data-stu-id="43f5c-331">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="43f5c-332">字符串</span><span class="sxs-lookup"><span data-stu-id="43f5c-332">String</span></span>|<span data-ttu-id="43f5c-333">现有日历约会的 Exchange Web 服务 (EWS) 标识符。</span><span class="sxs-lookup"><span data-stu-id="43f5c-333">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="43f5c-334">要求</span><span class="sxs-lookup"><span data-stu-id="43f5c-334">Requirements</span></span>

|<span data-ttu-id="43f5c-335">要求</span><span class="sxs-lookup"><span data-stu-id="43f5c-335">Requirement</span></span>| <span data-ttu-id="43f5c-336">值</span><span class="sxs-lookup"><span data-stu-id="43f5c-336">Value</span></span>|
|---|---|
|[<span data-ttu-id="43f5c-337">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="43f5c-337">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43f5c-338">1.0</span><span class="sxs-lookup"><span data-stu-id="43f5c-338">1.0</span></span>|
|[<span data-ttu-id="43f5c-339">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="43f5c-339">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43f5c-340">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43f5c-340">ReadItem</span></span>|
|[<span data-ttu-id="43f5c-341">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="43f5c-341">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="43f5c-342">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="43f5c-342">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="43f5c-343">示例</span><span class="sxs-lookup"><span data-stu-id="43f5c-343">Example</span></span>

```
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

####  <a name="displaymessageformitemid"></a><span data-ttu-id="43f5c-344">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="43f5c-344">displayMessageForm(itemId)</span></span>

<span data-ttu-id="43f5c-345">显示现有邮件。</span><span class="sxs-lookup"><span data-stu-id="43f5c-345">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="43f5c-346">在适用于 iOS 的 Outlook 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="43f5c-346">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="43f5c-347">`displayMessageForm` 方法将打开桌面新窗口中或移动设备对话框中的现有邮件。</span><span class="sxs-lookup"><span data-stu-id="43f5c-347">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="43f5c-348">在 Outlook Web App 中，此方法仅在窗体正文小于或等于 32 KB 字符数时，才会打开指定的窗体。</span><span class="sxs-lookup"><span data-stu-id="43f5c-348">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="43f5c-349">如果指定的项标识符未识别现有消息，则客户端计算机上不会显示任何消息，并且也不会返回错误消息。</span><span class="sxs-lookup"><span data-stu-id="43f5c-349">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="43f5c-p113">不要使用包含表示约会的 `itemId` 的 `displayMessageForm`。使用 `displayAppointmentForm` 方法显示现有的约会，并使用 `displayNewAppointmentForm` 显示窗体以新建约会。</span><span class="sxs-lookup"><span data-stu-id="43f5c-p113">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="43f5c-352">参数：</span><span class="sxs-lookup"><span data-stu-id="43f5c-352">Parameters:</span></span>

|<span data-ttu-id="43f5c-353">名称</span><span class="sxs-lookup"><span data-stu-id="43f5c-353">Name</span></span>| <span data-ttu-id="43f5c-354">类型</span><span class="sxs-lookup"><span data-stu-id="43f5c-354">Type</span></span>| <span data-ttu-id="43f5c-355">说明</span><span class="sxs-lookup"><span data-stu-id="43f5c-355">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="43f5c-356">字符串</span><span class="sxs-lookup"><span data-stu-id="43f5c-356">String</span></span>|<span data-ttu-id="43f5c-357">现有消息的 Exchange Web 服务 (EWS) 标识符。</span><span class="sxs-lookup"><span data-stu-id="43f5c-357">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="43f5c-358">要求</span><span class="sxs-lookup"><span data-stu-id="43f5c-358">Requirements</span></span>

|<span data-ttu-id="43f5c-359">要求</span><span class="sxs-lookup"><span data-stu-id="43f5c-359">Requirement</span></span>| <span data-ttu-id="43f5c-360">值</span><span class="sxs-lookup"><span data-stu-id="43f5c-360">Value</span></span>|
|---|---|
|[<span data-ttu-id="43f5c-361">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="43f5c-361">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43f5c-362">1.0</span><span class="sxs-lookup"><span data-stu-id="43f5c-362">1.0</span></span>|
|[<span data-ttu-id="43f5c-363">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="43f5c-363">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43f5c-364">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43f5c-364">ReadItem</span></span>|
|[<span data-ttu-id="43f5c-365">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="43f5c-365">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="43f5c-366">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="43f5c-366">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="43f5c-367">示例</span><span class="sxs-lookup"><span data-stu-id="43f5c-367">Example</span></span>

```
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="43f5c-368">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="43f5c-368">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="43f5c-369">显示用于新建日历约会的表单。</span><span class="sxs-lookup"><span data-stu-id="43f5c-369">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="43f5c-370">在适用于 iOS 的 Outlook 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="43f5c-370">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="43f5c-p114">`displayNewAppointmentForm` 方法打开可让用户新建约会或会议的窗体。如果指定了参数，将使用参数的内容自动填充约会窗体字段。</span><span class="sxs-lookup"><span data-stu-id="43f5c-p114">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="43f5c-p115">在 Outlook Web App 和适用于设备的 OWA 中，此方法始终显示包含与会者字段的窗体。如果你未将任何与会者指定为输入参数，该方法将显示为一个包含“**保存**”按钮的窗体。如果已指定与会者，窗体将包含与会者和“**发送**”按钮。</span><span class="sxs-lookup"><span data-stu-id="43f5c-p115">In Outlook Web App and OWA for Devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="43f5c-p116">在 Outlook 富客户端和 Outlook RT 中，如果在 `requiredAttendees`、`optionalAttendees` 或 `resources` 参数中指定任何与会者或资源，此方法将显示会议窗体，其中包含一个“**发送**”按钮。如果未指定任何收件人，此方法将显示一个包含“**保存并关闭**”按钮的约会窗体。</span><span class="sxs-lookup"><span data-stu-id="43f5c-p116">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="43f5c-378">如果任何参数超过指定大小限制，或者指定了未知参数名称，则会引发异常。</span><span class="sxs-lookup"><span data-stu-id="43f5c-378">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="43f5c-379">参数：</span><span class="sxs-lookup"><span data-stu-id="43f5c-379">Parameters:</span></span>

> [!NOTE]
> <span data-ttu-id="43f5c-380">所有参数都是可选的。</span><span class="sxs-lookup"><span data-stu-id="43f5c-380">All parameters are optional.</span></span>

|<span data-ttu-id="43f5c-381">名称</span><span class="sxs-lookup"><span data-stu-id="43f5c-381">Name</span></span>| <span data-ttu-id="43f5c-382">类型</span><span class="sxs-lookup"><span data-stu-id="43f5c-382">Type</span></span>| <span data-ttu-id="43f5c-383">说明</span><span class="sxs-lookup"><span data-stu-id="43f5c-383">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="43f5c-384">对象</span><span class="sxs-lookup"><span data-stu-id="43f5c-384">Object</span></span> | <span data-ttu-id="43f5c-385">描述新约会的参数字典。</span><span class="sxs-lookup"><span data-stu-id="43f5c-385">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="43f5c-386">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="43f5c-386">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="43f5c-p117">包含电子邮件地址的字符串数组或包含约会的每个必需与会者的 `EmailAddressDetails` 对象的数组。数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="43f5c-p117">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="43f5c-389">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="43f5c-389">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="43f5c-p118">包含电子邮件地址的字符串数组或包含约会的每个可选与会者的 `EmailAddressDetails` 对象的数组。数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="43f5c-p118">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="43f5c-392">Date</span><span class="sxs-lookup"><span data-stu-id="43f5c-392">Date</span></span> | <span data-ttu-id="43f5c-393">指定约会的开始日期和时间的 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="43f5c-393">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="43f5c-394">Date</span><span class="sxs-lookup"><span data-stu-id="43f5c-394">Date</span></span> | <span data-ttu-id="43f5c-395">指定约会的结束日期和时间的 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="43f5c-395">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="43f5c-396">String</span><span class="sxs-lookup"><span data-stu-id="43f5c-396">String</span></span> | <span data-ttu-id="43f5c-p119">包含约会位置的字符串。字符串长度限制为最多 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="43f5c-p119">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="43f5c-399">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="43f5c-399">Array.&lt;String&gt;</span></span> | <span data-ttu-id="43f5c-p120">包含约会所需资源的字符串数组。数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="43f5c-p120">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="43f5c-402">String</span><span class="sxs-lookup"><span data-stu-id="43f5c-402">String</span></span> | <span data-ttu-id="43f5c-p121">包含约会主题的字符串。字符串长度限制为最多 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="43f5c-p121">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="43f5c-405">String</span><span class="sxs-lookup"><span data-stu-id="43f5c-405">String</span></span> | <span data-ttu-id="43f5c-p122">约会的正文。正文内容限制为最大 32 KB。</span><span class="sxs-lookup"><span data-stu-id="43f5c-p122">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="43f5c-408">要求</span><span class="sxs-lookup"><span data-stu-id="43f5c-408">Requirements</span></span>

|<span data-ttu-id="43f5c-409">要求</span><span class="sxs-lookup"><span data-stu-id="43f5c-409">Requirement</span></span>| <span data-ttu-id="43f5c-410">值</span><span class="sxs-lookup"><span data-stu-id="43f5c-410">Value</span></span>|
|---|---|
|[<span data-ttu-id="43f5c-411">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="43f5c-411">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43f5c-412">1.0</span><span class="sxs-lookup"><span data-stu-id="43f5c-412">1.0</span></span>|
|[<span data-ttu-id="43f5c-413">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="43f5c-413">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43f5c-414">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43f5c-414">ReadItem</span></span>|
|[<span data-ttu-id="43f5c-415">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="43f5c-415">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="43f5c-416">阅读</span><span class="sxs-lookup"><span data-stu-id="43f5c-416">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="43f5c-417">示例</span><span class="sxs-lookup"><span data-stu-id="43f5c-417">Example</span></span>

```
var start = new Date();
var end = new Date();
end.setHours(start.getHours() + 1);

Office.context.mailbox.displayNewAppointmentForm(
  {
    requiredAttendees: ['bob@contoso.com'],
    optionalAttendees: ['sam@contoso.com'],
    start: start,
    end: end,
    location: 'Home',
    resources: ['projector@contoso.com'],
    subject: 'meeting',
    body: 'Hello World!'
  });
```

#### <a name="displaynewmessageformparameters"></a><span data-ttu-id="43f5c-418">displayNewMessageForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="43f5c-418">displayNewMessageForm(parameters)</span></span>

<span data-ttu-id="43f5c-419">显示用于新建邮件的窗体。</span><span class="sxs-lookup"><span data-stu-id="43f5c-419">Displays a form for creating a new message.</span></span>

<span data-ttu-id="43f5c-420">`displayNewMessageForm` 方法将打开可让用户新建邮件的窗体。</span><span class="sxs-lookup"><span data-stu-id="43f5c-420">The `displayNewMessageForm` method opens a form that enables the user to create a new message.</span></span> <span data-ttu-id="43f5c-421">如果指定了参数，将使用参数的内容自动填充邮件窗体字段。</span><span class="sxs-lookup"><span data-stu-id="43f5c-421">If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="43f5c-422">如果任何参数超过指定大小限制，或者指定了未知参数名称，则会引发异常。</span><span class="sxs-lookup"><span data-stu-id="43f5c-422">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="43f5c-423">参数：</span><span class="sxs-lookup"><span data-stu-id="43f5c-423">Parameters:</span></span>

> [!NOTE]
> <span data-ttu-id="43f5c-424">所有参数都是可选的。</span><span class="sxs-lookup"><span data-stu-id="43f5c-424">All parameters are optional.</span></span>

|<span data-ttu-id="43f5c-425">名称</span><span class="sxs-lookup"><span data-stu-id="43f5c-425">Name</span></span>| <span data-ttu-id="43f5c-426">类型</span><span class="sxs-lookup"><span data-stu-id="43f5c-426">Type</span></span>| <span data-ttu-id="43f5c-427">说明</span><span class="sxs-lookup"><span data-stu-id="43f5c-427">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="43f5c-428">对象</span><span class="sxs-lookup"><span data-stu-id="43f5c-428">Object</span></span> | <span data-ttu-id="43f5c-429">描述新邮件的参数字典。</span><span class="sxs-lookup"><span data-stu-id="43f5c-429">A dictionary of parameters describing the new message.</span></span> |
| `parameters.toRecipients` | <span data-ttu-id="43f5c-430">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="43f5c-430">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="43f5c-431">包含电子邮件地址的字符串数组或包含收件人行上每个收件人的 `EmailAddressDetails` 对象的数组。</span><span class="sxs-lookup"><span data-stu-id="43f5c-431">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line.</span></span> <span data-ttu-id="43f5c-432">数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="43f5c-432">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.ccRecipients` | <span data-ttu-id="43f5c-433">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="43f5c-433">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="43f5c-434">包含电子邮件地址的字符串数组或包含抄送行上每个收件人的 `EmailAddressDetails` 对象的数组。</span><span class="sxs-lookup"><span data-stu-id="43f5c-434">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line.</span></span> <span data-ttu-id="43f5c-435">数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="43f5c-435">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.bccRecipients` | <span data-ttu-id="43f5c-436">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="43f5c-436">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="43f5c-437">包含电子邮件地址的字符串数组或包含密件抄送行上每个收件人的 `EmailAddressDetails` 对象的数组。</span><span class="sxs-lookup"><span data-stu-id="43f5c-437">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line.</span></span> <span data-ttu-id="43f5c-438">数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="43f5c-438">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="43f5c-439">String</span><span class="sxs-lookup"><span data-stu-id="43f5c-439">String</span></span> | <span data-ttu-id="43f5c-440">包含邮件主题的字符串。</span><span class="sxs-lookup"><span data-stu-id="43f5c-440">A string containing the subject of the message.</span></span> <span data-ttu-id="43f5c-441">字符串长度限制为最多 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="43f5c-441">The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.htmlBody` | <span data-ttu-id="43f5c-442">String</span><span class="sxs-lookup"><span data-stu-id="43f5c-442">String</span></span> | <span data-ttu-id="43f5c-443">邮件的 HTML 正文。</span><span class="sxs-lookup"><span data-stu-id="43f5c-443">The HTML body of the message.</span></span> <span data-ttu-id="43f5c-444">正文内容限制为最大 32 KB。</span><span class="sxs-lookup"><span data-stu-id="43f5c-444">The body content is limited to a maximum size of 32 KB.</span></span> |
| `parameters.attachments` | <span data-ttu-id="43f5c-445">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="43f5c-445">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="43f5c-446">JSON 对象（文件或项目附件）数组。</span><span class="sxs-lookup"><span data-stu-id="43f5c-446">An array of JSON objects that are either file or item attachments.</span></span> |
| `parameters.attachments.type` | <span data-ttu-id="43f5c-447">字符串</span><span class="sxs-lookup"><span data-stu-id="43f5c-447">String</span></span> | <span data-ttu-id="43f5c-p129">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="43f5c-p129">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `parameters.attachments.name` | <span data-ttu-id="43f5c-450">字符串</span><span class="sxs-lookup"><span data-stu-id="43f5c-450">String</span></span> | <span data-ttu-id="43f5c-451">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="43f5c-451">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `parameters.attachments.url` | <span data-ttu-id="43f5c-452">String</span><span class="sxs-lookup"><span data-stu-id="43f5c-452">String</span></span> | <span data-ttu-id="43f5c-p130">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="43f5c-p130">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `parameters.attachments.isInline` | <span data-ttu-id="43f5c-455">Boolean</span><span class="sxs-lookup"><span data-stu-id="43f5c-455">Boolean</span></span> | <span data-ttu-id="43f5c-p131">仅在将 `type` 设置为 `file` 时使用。如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="43f5c-p131">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `parameters.attachments.itemId` | <span data-ttu-id="43f5c-458">String</span><span class="sxs-lookup"><span data-stu-id="43f5c-458">String</span></span> | <span data-ttu-id="43f5c-459">仅在 `type` 设置为 `item` 时使用。</span><span class="sxs-lookup"><span data-stu-id="43f5c-459">Only used if `type` is set to `item`.</span></span> <span data-ttu-id="43f5c-460">您想要附加到新邮件的现有电子邮件 EWS 项 id。</span><span class="sxs-lookup"><span data-stu-id="43f5c-460">The EWS item id of the existing e-mail you want to attach to the new message.</span></span> <span data-ttu-id="43f5c-461">字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="43f5c-461">This is a string up to 100 characters.</span></span> |


##### <a name="requirements"></a><span data-ttu-id="43f5c-462">要求</span><span class="sxs-lookup"><span data-stu-id="43f5c-462">Requirements</span></span>

|<span data-ttu-id="43f5c-463">要求</span><span class="sxs-lookup"><span data-stu-id="43f5c-463">Requirement</span></span>| <span data-ttu-id="43f5c-464">值</span><span class="sxs-lookup"><span data-stu-id="43f5c-464">Value</span></span>|
|---|---|
|[<span data-ttu-id="43f5c-465">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="43f5c-465">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43f5c-466">1.6</span><span class="sxs-lookup"><span data-stu-id="43f5c-466">1.6</span></span> |
|[<span data-ttu-id="43f5c-467">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="43f5c-467">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43f5c-468">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43f5c-468">ReadItem</span></span>|
|[<span data-ttu-id="43f5c-469">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="43f5c-469">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="43f5c-470">阅读</span><span class="sxs-lookup"><span data-stu-id="43f5c-470">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="43f5c-471">示例</span><span class="sxs-lookup"><span data-stu-id="43f5c-471">Example</span></span>

```
Office.context.mailbox.displayNewMessageForm(
  {
    toRecipients: Office.context.mailbox.item.to, // Copy the To line from current item
    ccRecipients: ['sam@contoso.com'],
    subject: 'Outlook add-ins are cool!',
    htmlBody: 'Hello <b>World</b>!<br/><img src="cid:image.png"></i>',
    attachments: [
      {
        type: 'file',
        name: 'image.png',
        url: 'http://contoso.com/image.png',
        isInline: true
      }
    ]
  });
```

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="43f5c-472">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="43f5c-472">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="43f5c-473">获取一个包含用于调用 REST API 或 Exchange Web 服务的令牌的字符串。</span><span class="sxs-lookup"><span data-stu-id="43f5c-473">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="43f5c-p133">`getCallbackTokenAsync` 方法进行异步调用，从托管用户邮箱的 Exchange Server 获取非跳转令牌。回调令牌的生存期为 5 分钟。</span><span class="sxs-lookup"><span data-stu-id="43f5c-p133">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="43f5c-476">建议外接程序而不是 Exchange Web 服务尽可能使用 REST Api。</span><span class="sxs-lookup"><span data-stu-id="43f5c-476">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span> 

<span data-ttu-id="43f5c-477">**REST 令牌**</span><span class="sxs-lookup"><span data-stu-id="43f5c-477">**REST Tokens**</span></span>

<span data-ttu-id="43f5c-p134">请求 REST 令牌时 (`options.isRest = true`) 时，生成的令牌将无法对 Exchange Web 服务调用进行身份验证。令牌的作用域限制为对当前项及其附件的只读访问，除非外接程序在其清单中指定了 [`ReadWriteMailbox`](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) 权限。如果指定了 `ReadWriteMailbox` 权限，则生成的令牌将授予对邮件、日历和联系人的读/写权限，包括发送邮件的功能。</span><span class="sxs-lookup"><span data-stu-id="43f5c-p134">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="43f5c-481">在进行 REST API 调用时，外接程序应使用 `restUrl` 属性来确定要使用的正确 URL。</span><span class="sxs-lookup"><span data-stu-id="43f5c-481">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="43f5c-482">**EWS 令牌**</span><span class="sxs-lookup"><span data-stu-id="43f5c-482">**EWS Tokens**</span></span>

<span data-ttu-id="43f5c-p135">请求 EWS 令牌 (`options.isRest = false`) 时，生成的令牌将无法对 REST API 调用进行身份验证。令牌的作用域限制为访问当前项。</span><span class="sxs-lookup"><span data-stu-id="43f5c-p135">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="43f5c-485">外接程序应使用 `ewsUrl` 属性来确定进行 EWS 调用时要使用的正确 URL。</span><span class="sxs-lookup"><span data-stu-id="43f5c-485">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

##### <a name="parameters"></a><span data-ttu-id="43f5c-486">参数：</span><span class="sxs-lookup"><span data-stu-id="43f5c-486">Parameters:</span></span>

|<span data-ttu-id="43f5c-487">名称</span><span class="sxs-lookup"><span data-stu-id="43f5c-487">Name</span></span>| <span data-ttu-id="43f5c-488">类型</span><span class="sxs-lookup"><span data-stu-id="43f5c-488">Type</span></span>| <span data-ttu-id="43f5c-489">属性</span><span class="sxs-lookup"><span data-stu-id="43f5c-489">Attributes</span></span>| <span data-ttu-id="43f5c-490">说明</span><span class="sxs-lookup"><span data-stu-id="43f5c-490">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="43f5c-491">对象</span><span class="sxs-lookup"><span data-stu-id="43f5c-491">Object</span></span> | <span data-ttu-id="43f5c-492">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="43f5c-492">&lt;optional&gt;</span></span> | <span data-ttu-id="43f5c-493">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="43f5c-493">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="43f5c-494">布尔值</span><span class="sxs-lookup"><span data-stu-id="43f5c-494">Boolean</span></span> |  <span data-ttu-id="43f5c-495">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="43f5c-495">&lt;optional&gt;</span></span> | <span data-ttu-id="43f5c-p136">确定所提供的令牌是否将用于 Outlook REST API 或 Exchange Web 服务。默认值为 `false`。</span><span class="sxs-lookup"><span data-stu-id="43f5c-p136">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="43f5c-498">对象</span><span class="sxs-lookup"><span data-stu-id="43f5c-498">Object</span></span> |  <span data-ttu-id="43f5c-499">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="43f5c-499">&lt;optional&gt;</span></span> | <span data-ttu-id="43f5c-500">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="43f5c-500">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="43f5c-501">函数</span><span class="sxs-lookup"><span data-stu-id="43f5c-501">function</span></span>||<span data-ttu-id="43f5c-p137">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。令牌作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="43f5c-p137">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="43f5c-504">要求</span><span class="sxs-lookup"><span data-stu-id="43f5c-504">Requirements</span></span>

|<span data-ttu-id="43f5c-505">要求</span><span class="sxs-lookup"><span data-stu-id="43f5c-505">Requirement</span></span>| <span data-ttu-id="43f5c-506">值</span><span class="sxs-lookup"><span data-stu-id="43f5c-506">Value</span></span>|
|---|---|
|[<span data-ttu-id="43f5c-507">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="43f5c-507">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43f5c-508">1.5</span><span class="sxs-lookup"><span data-stu-id="43f5c-508">1.5</span></span> |
|[<span data-ttu-id="43f5c-509">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="43f5c-509">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43f5c-510">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43f5c-510">ReadItem</span></span>|
|[<span data-ttu-id="43f5c-511">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="43f5c-511">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="43f5c-512">撰写和阅读</span><span class="sxs-lookup"><span data-stu-id="43f5c-512">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="43f5c-513">示例</span><span class="sxs-lookup"><span data-stu-id="43f5c-513">Example</span></span>

```js
function getCallbackToken() {
  var options = {
    isRest: true,
    asyncContext: { message: 'Hello World!' }
  };

  Office.context.mailbox.getCallbackTokenAsync(options, cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="43f5c-514">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="43f5c-514">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="43f5c-515">获取一个字符串，其中包含用于从 Exchange Server 获取附件或项目的令牌。</span><span class="sxs-lookup"><span data-stu-id="43f5c-515">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="43f5c-p138">`getCallbackTokenAsync` 方法进行异步调用，从托管用户邮箱的 Exchange Server 获取非跳转令牌。回调令牌的生存期为 5 分钟。</span><span class="sxs-lookup"><span data-stu-id="43f5c-p138">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="43f5c-p139">可以将令牌和附件标识符或项标识符传递到第三方系统。第三方系统使用令牌作为持有者身份验证令牌调用 Exchange Web 服务 (EWS) [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) 或 [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation)，以返回附件或项目。例如，可以创建远程服务来[获取选定项目中的附件](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item)。</span><span class="sxs-lookup"><span data-stu-id="43f5c-p139">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="43f5c-521">应用必须在其清单中指定拥有 **ReadItem** 权限，才能调用阅读模式中的 `getCallbackTokenAsync` 方法。</span><span class="sxs-lookup"><span data-stu-id="43f5c-521">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="43f5c-p140">在撰写模式中，必须调用 [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) 方法来获取传递给 `getCallbackTokenAsync` 方法的项目标识符。应用必须具有调用 `saveAsync` 方法的 **ReadWriteItem** 权限。</span><span class="sxs-lookup"><span data-stu-id="43f5c-p140">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="43f5c-524">参数：</span><span class="sxs-lookup"><span data-stu-id="43f5c-524">Parameters:</span></span>

|<span data-ttu-id="43f5c-525">名称</span><span class="sxs-lookup"><span data-stu-id="43f5c-525">Name</span></span>| <span data-ttu-id="43f5c-526">类型</span><span class="sxs-lookup"><span data-stu-id="43f5c-526">Type</span></span>| <span data-ttu-id="43f5c-527">属性</span><span class="sxs-lookup"><span data-stu-id="43f5c-527">Attributes</span></span>| <span data-ttu-id="43f5c-528">说明</span><span class="sxs-lookup"><span data-stu-id="43f5c-528">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="43f5c-529">函数</span><span class="sxs-lookup"><span data-stu-id="43f5c-529">function</span></span>||<span data-ttu-id="43f5c-p141">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。令牌作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="43f5c-p141">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="43f5c-532">对象</span><span class="sxs-lookup"><span data-stu-id="43f5c-532">Object</span></span>| <span data-ttu-id="43f5c-533">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="43f5c-533">&lt;optional&gt;</span></span>|<span data-ttu-id="43f5c-534">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="43f5c-534">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="43f5c-535">要求</span><span class="sxs-lookup"><span data-stu-id="43f5c-535">Requirements</span></span>

|<span data-ttu-id="43f5c-536">要求</span><span class="sxs-lookup"><span data-stu-id="43f5c-536">Requirement</span></span>| <span data-ttu-id="43f5c-537">值</span><span class="sxs-lookup"><span data-stu-id="43f5c-537">Value</span></span>|
|---|---|
|[<span data-ttu-id="43f5c-538">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="43f5c-538">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43f5c-539">1.3</span><span class="sxs-lookup"><span data-stu-id="43f5c-539">1.3</span></span>|
|[<span data-ttu-id="43f5c-540">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="43f5c-540">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43f5c-541">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43f5c-541">ReadItem</span></span>|
|[<span data-ttu-id="43f5c-542">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="43f5c-542">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="43f5c-543">撰写和阅读</span><span class="sxs-lookup"><span data-stu-id="43f5c-543">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="43f5c-544">示例</span><span class="sxs-lookup"><span data-stu-id="43f5c-544">Example</span></span>

```js
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="43f5c-545">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="43f5c-545">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="43f5c-546">获取用于标识用户和 Office 外接程序的令牌。</span><span class="sxs-lookup"><span data-stu-id="43f5c-546">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="43f5c-547">`getUserIdentityTokenAsync` 方法返回你可以用于在第三方系统上识别和 [验证外接程序和用户的令牌](https://docs.microsoft.com/outlook/add-ins/authentication)。</span><span class="sxs-lookup"><span data-stu-id="43f5c-547">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](https://docs.microsoft.com/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="43f5c-548">参数：</span><span class="sxs-lookup"><span data-stu-id="43f5c-548">Parameters:</span></span>

|<span data-ttu-id="43f5c-549">名称</span><span class="sxs-lookup"><span data-stu-id="43f5c-549">Name</span></span>| <span data-ttu-id="43f5c-550">类型</span><span class="sxs-lookup"><span data-stu-id="43f5c-550">Type</span></span>| <span data-ttu-id="43f5c-551">属性</span><span class="sxs-lookup"><span data-stu-id="43f5c-551">Attributes</span></span>| <span data-ttu-id="43f5c-552">说明</span><span class="sxs-lookup"><span data-stu-id="43f5c-552">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="43f5c-553">函数</span><span class="sxs-lookup"><span data-stu-id="43f5c-553">function</span></span>||<span data-ttu-id="43f5c-554">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="43f5c-554">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="43f5c-555">令牌作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="43f5c-555">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="43f5c-556">对象</span><span class="sxs-lookup"><span data-stu-id="43f5c-556">Object</span></span>| <span data-ttu-id="43f5c-557">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="43f5c-557">&lt;optional&gt;</span></span>|<span data-ttu-id="43f5c-558">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="43f5c-558">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="43f5c-559">要求</span><span class="sxs-lookup"><span data-stu-id="43f5c-559">Requirements</span></span>

|<span data-ttu-id="43f5c-560">要求</span><span class="sxs-lookup"><span data-stu-id="43f5c-560">Requirement</span></span>| <span data-ttu-id="43f5c-561">值</span><span class="sxs-lookup"><span data-stu-id="43f5c-561">Value</span></span>|
|---|---|
|[<span data-ttu-id="43f5c-562">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="43f5c-562">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43f5c-563">1.0</span><span class="sxs-lookup"><span data-stu-id="43f5c-563">1.0</span></span>|
|[<span data-ttu-id="43f5c-564">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="43f5c-564">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43f5c-565">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43f5c-565">ReadItem</span></span>|
|[<span data-ttu-id="43f5c-566">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="43f5c-566">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="43f5c-567">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="43f5c-567">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="43f5c-568">示例</span><span class="sxs-lookup"><span data-stu-id="43f5c-568">Example</span></span>

```js
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="43f5c-569">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="43f5c-569">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="43f5c-570">向托管用户邮箱的 Exchange 服务器上的 Exchange Web 服务 (EWS) 发出异步请求。</span><span class="sxs-lookup"><span data-stu-id="43f5c-570">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="43f5c-571">在以下方案中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="43f5c-571">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="43f5c-572">在适用于 iOS 的 Outlook 或 Outlook for Android</span><span class="sxs-lookup"><span data-stu-id="43f5c-572">In Outlook for iOS or Outlook for Android</span></span>
> - <span data-ttu-id="43f5c-573">外接程序加载时在 Gmail 邮箱</span><span class="sxs-lookup"><span data-stu-id="43f5c-573">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="43f5c-574">在这些情况下外, 接程序应[使用 REST Api](https://docs.microsoft.com/outlook/add-ins/use-rest-api)访问用户的邮箱，而是。</span><span class="sxs-lookup"><span data-stu-id="43f5c-574">In these cases, add-ins should [use REST APIs](https://docs.microsoft.com/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="43f5c-575">`makeEwsRequestAsync` 方法代表外接程序将 EWS 请求发送到 Exchange。</span><span class="sxs-lookup"><span data-stu-id="43f5c-575">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="43f5c-576">有关支持的 EWS 操作的列表，请参阅[调用 web 服务从 Outlook 外接程序](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support)。</span><span class="sxs-lookup"><span data-stu-id="43f5c-576">See [Call web services from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="43f5c-577">你不能使用 `makeEwsRequestAsync` 方法请求与文件夹关联的项目。</span><span class="sxs-lookup"><span data-stu-id="43f5c-577">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="43f5c-578">XML 请求必须指定 UTF-8 编码。</span><span class="sxs-lookup"><span data-stu-id="43f5c-578">The XML request must specify UTF-8 encoding.</span></span>

```
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="43f5c-p143">外接程序必须具有 **ReadWriteMailbox** 权限才能使用 `makeEwsRequestAsync` 方法。有关使用 **ReadWriteMailbox** 权限和可使用 `makeEwsRequestAsync` 方法调用 EWS 操作的信息，请参阅[指定访问用户邮箱的邮件外接程序的权限](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)。</span><span class="sxs-lookup"><span data-stu-id="43f5c-p143">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="43f5c-581">服务器管理员必须设置`OAuthAuthentication`为 true 时，要启用的客户端访问服务器 EWS 目录上`makeEwsRequestAsync`方法来发出 EWS 请求。</span><span class="sxs-lookup"><span data-stu-id="43f5c-581">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="43f5c-582">版本差异</span><span class="sxs-lookup"><span data-stu-id="43f5c-582">Version differences</span></span>

<span data-ttu-id="43f5c-583">当你在较 15.0.4535.1004 版本更早的 Outlook 版本中运行的邮件应用程序中使用 `makeEwsRequestAsync` 方法，应当将编码值设置为 `ISO-8859-1`。</span><span class="sxs-lookup"><span data-stu-id="43f5c-583">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="43f5c-p144">当邮件应用程序运行在 Outlook 网页版中时，您不需要设置编码值。可以通过使用 mailbox.diagnostics.hostName 属性来确定您的邮件应用程序在 Outlook 中还是 Outlook 网页版中运行。可以通过使用 mailbox.diagnostics.hostVersion 属性来确定正在运行的是 Outlook 的哪个版本。</span><span class="sxs-lookup"><span data-stu-id="43f5c-p144">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="43f5c-587">参数：</span><span class="sxs-lookup"><span data-stu-id="43f5c-587">Parameters:</span></span>

|<span data-ttu-id="43f5c-588">名称</span><span class="sxs-lookup"><span data-stu-id="43f5c-588">Name</span></span>| <span data-ttu-id="43f5c-589">类型</span><span class="sxs-lookup"><span data-stu-id="43f5c-589">Type</span></span>| <span data-ttu-id="43f5c-590">属性</span><span class="sxs-lookup"><span data-stu-id="43f5c-590">Attributes</span></span>| <span data-ttu-id="43f5c-591">说明</span><span class="sxs-lookup"><span data-stu-id="43f5c-591">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="43f5c-592">字符串</span><span class="sxs-lookup"><span data-stu-id="43f5c-592">String</span></span>||<span data-ttu-id="43f5c-593">EWS 请求。</span><span class="sxs-lookup"><span data-stu-id="43f5c-593">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="43f5c-594">函数</span><span class="sxs-lookup"><span data-stu-id="43f5c-594">function</span></span>||<span data-ttu-id="43f5c-595">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="43f5c-595">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="43f5c-596">EWS 调用的 XML 结果作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="43f5c-596">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="43f5c-597">如果结果的大小超过 1 MB，而被返回一条错误消息。</span><span class="sxs-lookup"><span data-stu-id="43f5c-597">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="43f5c-598">对象</span><span class="sxs-lookup"><span data-stu-id="43f5c-598">Object</span></span>| <span data-ttu-id="43f5c-599">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="43f5c-599">&lt;optional&gt;</span></span>|<span data-ttu-id="43f5c-600">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="43f5c-600">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="43f5c-601">要求</span><span class="sxs-lookup"><span data-stu-id="43f5c-601">Requirements</span></span>

|<span data-ttu-id="43f5c-602">要求</span><span class="sxs-lookup"><span data-stu-id="43f5c-602">Requirement</span></span>| <span data-ttu-id="43f5c-603">值</span><span class="sxs-lookup"><span data-stu-id="43f5c-603">Value</span></span>|
|---|---|
|[<span data-ttu-id="43f5c-604">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="43f5c-604">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43f5c-605">1.0</span><span class="sxs-lookup"><span data-stu-id="43f5c-605">1.0</span></span>|
|[<span data-ttu-id="43f5c-606">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="43f5c-606">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43f5c-607">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="43f5c-607">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="43f5c-608">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="43f5c-608">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="43f5c-609">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="43f5c-609">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="43f5c-610">示例</span><span class="sxs-lookup"><span data-stu-id="43f5c-610">Example</span></span>

<span data-ttu-id="43f5c-611">下面的示例调用 `makeEwsRequestAsync` 以使用 `GetItem` 操作来获取项目的主题。</span><span class="sxs-lookup"><span data-stu-id="43f5c-611">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

```js
function getSubjectRequest(id) {
   // Return a GetItem operation request for the subject of the specified item.
   var request =
    '<?xml version="1.0" encoding="utf-8"?>' +
    '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
    '               xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
    '               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
    '               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
    '  <soap:Header>' +
    '    <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
    '  </soap:Header>' +
    '  <soap:Body>' +
    '    <GetItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">' +
    '      <ItemShape>' +
    '        <t:BaseShape>IdOnly</t:BaseShape>' +
    '        <t:AdditionalProperties>' +
    '            <t:FieldURI FieldURI="item:Subject"/>' +
    '        </t:AdditionalProperties>' +
    '      </ItemShape>' +
    '      <ItemIds><t:ItemId Id="' + id + '"/></ItemIds>' +
    '    </GetItem>' +
    '  </soap:Body>' +
    '</soap:Envelope>';

   return request;
}

function sendRequest() {
   // Create a local variable that contains the mailbox.
   Office.context.mailbox.makeEwsRequestAsync(
    getSubjectRequest(mailbox.item.itemId), callback);
}

function callback(asyncResult)  {
   var result = asyncResult.value;
   var context = asyncResult.asyncContext;

   // Process the returned response here.
}
```