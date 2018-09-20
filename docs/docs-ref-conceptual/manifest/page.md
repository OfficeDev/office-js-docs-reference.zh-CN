# <a name="page-element"></a><span data-ttu-id="a2b2d-101">Page 元素</span><span class="sxs-lookup"><span data-stu-id="a2b2d-101">Page element</span></span>

<span data-ttu-id="a2b2d-102">定义 Excel 中的自定义函数所使用的 HTML 页面设置。</span><span class="sxs-lookup"><span data-stu-id="a2b2d-102">Defines HTML page settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="a2b2d-103">属性</span><span class="sxs-lookup"><span data-stu-id="a2b2d-103">Attributes</span></span>

<span data-ttu-id="a2b2d-104">无</span><span class="sxs-lookup"><span data-stu-id="a2b2d-104">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="a2b2d-105">子元素</span><span class="sxs-lookup"><span data-stu-id="a2b2d-105">Child elements</span></span>

|  <span data-ttu-id="a2b2d-106">元素</span><span class="sxs-lookup"><span data-stu-id="a2b2d-106">Element</span></span>  |  <span data-ttu-id="a2b2d-107">必需</span><span class="sxs-lookup"><span data-stu-id="a2b2d-107">Required</span></span>  |  <span data-ttu-id="a2b2d-108">说明</span><span class="sxs-lookup"><span data-stu-id="a2b2d-108">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="a2b2d-109">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="a2b2d-109">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="a2b2d-110">是</span><span class="sxs-lookup"><span data-stu-id="a2b2d-110">Yes</span></span>  | <span data-ttu-id="a2b2d-111">包含自定义函数所使用的 HTML 文件的资源 ID 的字符串。</span><span class="sxs-lookup"><span data-stu-id="a2b2d-111">String with the resource id of the HTML file used by custom functions.</span></span> |

## <a name="example"></a><span data-ttu-id="a2b2d-112">示例</span><span class="sxs-lookup"><span data-stu-id="a2b2d-112">Example</span></span>

```xml
<Page>
    <SourceLocation resid="pageURL"/>
</Page>
```
