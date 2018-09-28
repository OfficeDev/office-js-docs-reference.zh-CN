# <a name="outlook-add-in-api-requirement-set-12"></a><span data-ttu-id="7c585-101">Outlook 外接程序 API 要求集 1.2</span><span class="sxs-lookup"><span data-stu-id="7c585-101">Outlook add-in API requirement set 1.2</span></span>

<span data-ttu-id="7c585-102">Outlook 加载项 API 子集的 JavaScript API for Office 包括对象、 方法、 属性和事件，您可以在 Outlook 中使用外接程序。</span><span class="sxs-lookup"><span data-stu-id="7c585-102">The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="7c585-103">本文档是[要求设置](/javascript/office/requirement-sets/outlook-api-requirement-sets)而不是最新的要求集。</span><span class="sxs-lookup"><span data-stu-id="7c585-103">This documentation is for a [requirement set](/javascript/office/requirement-sets/outlook-api-requirement-sets) other than the latest requirement set.</span></span> 

## <a name="whats-new-in-12"></a><span data-ttu-id="7c585-104">1.2 中的新增功能有哪些？</span><span class="sxs-lookup"><span data-stu-id="7c585-104">What's new in 1.2?</span></span>

<span data-ttu-id="7c585-p101">要求集 1.2 包括[要求集 1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) 的所有功能。它添加了外接程序在用户的游标中插入文本的功能，无论是在邮件的主题还是正文中皆可插入文本。</span><span class="sxs-lookup"><span data-stu-id="7c585-p101">Requirement set 1.2 includes all of the features of [Requirement set 1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md). It added the ability for add-ins to insert text at the user's cursor, either in the subject or the body of the message.</span></span>

### <a name="change-log"></a><span data-ttu-id="7c585-107">更改日志</span><span class="sxs-lookup"><span data-stu-id="7c585-107">Change log</span></span>

- <span data-ttu-id="7c585-108">添加[Office.context.mailbox.item.getSelectedDataAsync](office.context.mailbox.item.md#getselecteddataasynccoerciontype-options-callback--string)： 中的主题或邮件正文的选定的数据的异步返回。</span><span class="sxs-lookup"><span data-stu-id="7c585-108">Added [Office.context.mailbox.item.getSelectedDataAsync](office.context.mailbox.item.md#getselecteddataasynccoerciontype-options-callback--string): Asynchronously returns selected data from the subject or body of a message.</span></span>
- <span data-ttu-id="7c585-109">添加了 [Office.context.mailbox.item.setSelectedDataAsync](office.context.mailbox.item.md#setselecteddataasyncdata-options-callback)：以异步方式将数据插入到邮件的正文或主题中。</span><span class="sxs-lookup"><span data-stu-id="7c585-109">Added [Office.context.mailbox.item.setSelectedDataAsync](office.context.mailbox.item.md#setselecteddataasyncdata-options-callback): Asynchronously inserts data into the body or subject of a message.</span></span>
- <span data-ttu-id="7c585-110">修改了 [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#displayreplyallformformdata)：将 `attachments` 属性添加到 `formData` 参数。</span><span class="sxs-lookup"><span data-stu-id="7c585-110">Modified [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#displayreplyallformformdata): Added `attachments` property to the `formData` parameter.</span></span>
- <span data-ttu-id="7c585-111">修改了 [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#displayreplyformformdata)：将 `attachments` 属性添加到 `formData` 参数。</span><span class="sxs-lookup"><span data-stu-id="7c585-111">Modified [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#displayreplyformformdata): Added `attachments` property to the `formData` parameter.</span></span>

## <a name="see-also"></a><span data-ttu-id="7c585-112">另请参阅</span><span class="sxs-lookup"><span data-stu-id="7c585-112">See also</span></span>

- [<span data-ttu-id="7c585-113">Outlook 外接程序</span><span class="sxs-lookup"><span data-stu-id="7c585-113">Outlook add-ins</span></span>](https://docs.microsoft.com/outlook/add-ins/)
- [<span data-ttu-id="7c585-114">Outlook 外接程序代码示例</span><span class="sxs-lookup"><span data-stu-id="7c585-114">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="7c585-115">入门</span><span class="sxs-lookup"><span data-stu-id="7c585-115">Get started</span></span>](https://docs.microsoft.com/outlook/add-ins/quick-start)