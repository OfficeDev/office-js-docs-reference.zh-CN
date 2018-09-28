# <a name="outlook-add-in-api-requirement-set-14"></a><span data-ttu-id="21b1f-101">Outlook 外接程序 API 要求集 1.4</span><span class="sxs-lookup"><span data-stu-id="21b1f-101">Outlook add-in API requirement set 1.4</span></span>

<span data-ttu-id="21b1f-102">Outlook 加载项 API 子集的 JavaScript API for Office 包括对象、 方法、 属性和事件，您可以在 Outlook 中使用外接程序。</span><span class="sxs-lookup"><span data-stu-id="21b1f-102">The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="21b1f-103">本文档是[要求设置](/javascript/office/requirement-sets/outlook-api-requirement-sets)而不是最新的要求集。</span><span class="sxs-lookup"><span data-stu-id="21b1f-103">This documentation is for a [requirement set](/javascript/office/requirement-sets/outlook-api-requirement-sets) other than the latest requirement set.</span></span>

## <a name="whats-new-in-14"></a><span data-ttu-id="21b1f-104">1.4 中的新增功能有哪些？</span><span class="sxs-lookup"><span data-stu-id="21b1f-104">What's new in 1.4?</span></span>

<span data-ttu-id="21b1f-p101">要求集 1.4 包括[要求集 1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) 的所有功能。它添加了对 `Office.ui` 命名空间的访问权限。</span><span class="sxs-lookup"><span data-stu-id="21b1f-p101">Requirement set 1.4 includes all of the features of [Requirement set 1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md). It added access to the `Office.ui` namespace.</span></span>

### <a name="change-log"></a><span data-ttu-id="21b1f-107">更改日志</span><span class="sxs-lookup"><span data-stu-id="21b1f-107">Change log</span></span>

- <span data-ttu-id="21b1f-108">添加了 [Office.context.ui.displayDialogAsync](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-)：在 Office 主机中显示一个对话框。</span><span class="sxs-lookup"><span data-stu-id="21b1f-108">Added [Office.context.ui.displayDialogAsync](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-): Displays a dialog box in an Office host.</span></span>
- <span data-ttu-id="21b1f-109">添加了 [Office.context.ui.messageParent](/javascript/api/office/office.ui#messageparent-messageobject-)：将对话框中的消息传送到其父页/开始页。</span><span class="sxs-lookup"><span data-stu-id="21b1f-109">Added [Office.context.ui.messageParent](/javascript/api/office/office.ui#messageparent-messageobject-): Delivers a message from the dialog box to its parent/opener page.</span></span>
- <span data-ttu-id="21b1f-110">添加了的[Dialog](/javascript/api/office/office.dialog)对象： 时返回的对象[`displayDialogAsync`](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-)调用方法。</span><span class="sxs-lookup"><span data-stu-id="21b1f-110">Added [Dialog](/javascript/api/office/office.dialog) object: The object that is returned when the [`displayDialogAsync`](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-) method is called.</span></span>

## <a name="see-also"></a><span data-ttu-id="21b1f-111">另请参阅</span><span class="sxs-lookup"><span data-stu-id="21b1f-111">See also</span></span>

- [<span data-ttu-id="21b1f-112">Outlook 外接程序</span><span class="sxs-lookup"><span data-stu-id="21b1f-112">Outlook add-ins</span></span>](https://docs.microsoft.com/outlook/add-ins/)
- [<span data-ttu-id="21b1f-113">Outlook 外接程序代码示例</span><span class="sxs-lookup"><span data-stu-id="21b1f-113">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="21b1f-114">入门</span><span class="sxs-lookup"><span data-stu-id="21b1f-114">Get started</span></span>](https://docs.microsoft.com/outlook/add-ins/quick-start)