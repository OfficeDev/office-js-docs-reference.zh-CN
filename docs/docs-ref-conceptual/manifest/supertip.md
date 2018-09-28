# <a name="supertip"></a><span data-ttu-id="a9b69-101">Supertip</span><span class="sxs-lookup"><span data-stu-id="a9b69-101">Supertip</span></span>

<span data-ttu-id="a9b69-p101">定义丰富的工具提示（标题和说明）。它由“[按钮](control.md#button-control)”或“[菜单](control.md#menu-dropdown-button-controls)”控件使用。</span><span class="sxs-lookup"><span data-stu-id="a9b69-p101">Defines a rich tooltip (both Title and Description). It is used by both [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls)  controls.</span></span>

## <a name="child-elements"></a><span data-ttu-id="a9b69-104">子元素</span><span class="sxs-lookup"><span data-stu-id="a9b69-104">Child elements</span></span>

|  <span data-ttu-id="a9b69-105">元素</span><span class="sxs-lookup"><span data-stu-id="a9b69-105">Element</span></span> |  <span data-ttu-id="a9b69-106">必需</span><span class="sxs-lookup"><span data-stu-id="a9b69-106">Required</span></span>  |  <span data-ttu-id="a9b69-107">说明</span><span class="sxs-lookup"><span data-stu-id="a9b69-107">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="a9b69-108">标题</span><span class="sxs-lookup"><span data-stu-id="a9b69-108">Title</span></span>](#title)        | <span data-ttu-id="a9b69-109">是</span><span class="sxs-lookup"><span data-stu-id="a9b69-109">Yes</span></span> |   <span data-ttu-id="a9b69-110">supertip 的文本。</span><span class="sxs-lookup"><span data-stu-id="a9b69-110">The text for the supertip.</span></span>         |
|  [<span data-ttu-id="a9b69-111">说明</span><span class="sxs-lookup"><span data-stu-id="a9b69-111">Description</span></span>](#description)  | <span data-ttu-id="a9b69-112">是</span><span class="sxs-lookup"><span data-stu-id="a9b69-112">Yes</span></span> |  <span data-ttu-id="a9b69-113">supertip 的说明。</span><span class="sxs-lookup"><span data-stu-id="a9b69-113">The description for the supertip.</span></span>    |

### <a name="title"></a><span data-ttu-id="a9b69-114">标题</span><span class="sxs-lookup"><span data-stu-id="a9b69-114">Title</span></span>

<span data-ttu-id="a9b69-p102">必需。SuperTip 的文本。 **resid** 属性必须设置为 **ShortStrings** 元素（位于 **Resources** 元素）中 **String** 元素的 [id](resources.md) 属性的值。</span><span class="sxs-lookup"><span data-stu-id="a9b69-p102">Required. The text for the supertip. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="description"></a><span data-ttu-id="a9b69-118">说明</span><span class="sxs-lookup"><span data-stu-id="a9b69-118">Description</span></span>

<span data-ttu-id="a9b69-p103">必需。SuperTip 的描述。 **resid** 属性必须设置为 **LongStrings** 元素（位于 **Resources** 元素）中 **String** 元素的 [id](resources.md) 属性的值。</span><span class="sxs-lookup"><span data-stu-id="a9b69-p103">Required. The description for the supertip. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **LongStrings** element in the [Resources](resources.md) element.</span></span>

## <a name="example"></a><span data-ttu-id="a9b69-122">示例</span><span class="sxs-lookup"><span data-stu-id="a9b69-122">Example</span></span>

```xml
<Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
</Supertip>
```
