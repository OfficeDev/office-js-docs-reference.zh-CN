# <a name="metadata-element"></a><span data-ttu-id="aa6e4-101">Metadata 元素</span><span class="sxs-lookup"><span data-stu-id="aa6e4-101">Metadata element</span></span>

<span data-ttu-id="aa6e4-102">定义自定义函数在 Excel 中使用的元数据设置。</span><span class="sxs-lookup"><span data-stu-id="aa6e4-102">Defines the metadata settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="aa6e4-103">属性</span><span class="sxs-lookup"><span data-stu-id="aa6e4-103">Attributes</span></span>

<span data-ttu-id="aa6e4-104">无</span><span class="sxs-lookup"><span data-stu-id="aa6e4-104">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="aa6e4-105">子元素</span><span class="sxs-lookup"><span data-stu-id="aa6e4-105">Child elements</span></span>

|  <span data-ttu-id="aa6e4-106">元素</span><span class="sxs-lookup"><span data-stu-id="aa6e4-106">Element</span></span>  |  <span data-ttu-id="aa6e4-107">必需</span><span class="sxs-lookup"><span data-stu-id="aa6e4-107">Required</span></span>  |  <span data-ttu-id="aa6e4-108">说明</span><span class="sxs-lookup"><span data-stu-id="aa6e4-108">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="aa6e4-109">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="aa6e4-109">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="aa6e4-110">是</span><span class="sxs-lookup"><span data-stu-id="aa6e4-110">Yes</span></span>  | <span data-ttu-id="aa6e4-111">使用自定义函数的 JSON 文件的资源 id 的字符串。</span><span class="sxs-lookup"><span data-stu-id="aa6e4-111">String with the resource id of the JSON file used by custom functions.</span></span> |

## <a name="example"></a><span data-ttu-id="aa6e4-112">示例</span><span class="sxs-lookup"><span data-stu-id="aa6e4-112">Example</span></span>

```xml
<Metadata>
    <SourceLocation resid="JSON-URL" />
</Metadata>
```
