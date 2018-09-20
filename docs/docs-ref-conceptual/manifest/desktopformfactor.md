# <a name="desktopformfactor-element"></a><span data-ttu-id="591e1-101">DesktopFormFactor 元素</span><span class="sxs-lookup"><span data-stu-id="591e1-101">DesktopFormFactor element</span></span>

<span data-ttu-id="591e1-p101">指定对桌面外形规格的外接程序的设置。桌面外形规格包括 Office for Windows、Office for Mac 和 Office Online。它包含该外形规格的所有外接程序信息（**资源**节点的信息除外）。</span><span class="sxs-lookup"><span data-stu-id="591e1-p101">Specifies the settings for an add-in for the desktop form factor. The desktop form factor includes Office for Windows, Office for Mac, and Office Online. It contains all the add-in information for the desktop form factor except for the  **Resources** node.</span></span>

<span data-ttu-id="591e1-p102">每个 DesktopFormFactor 定义均包含 **FunctionFile** 元素和一个或多个 **ExtensionPoint** 元素。有关详细信息，请参阅 [FunctionFile 元素](functionfile.md) 和 [ExtensionPoint 元素](extensionpoint.md)。</span><span class="sxs-lookup"><span data-stu-id="591e1-p102">Each DesktopFormFactor definition contains the  **FunctionFile** element and one or more **ExtensionPoint** elements. For more information, see [FunctionFile element](functionfile.md) and [ExtensionPoint element](extensionpoint.md).</span></span> 

## <a name="child-elements"></a><span data-ttu-id="591e1-107">子元素</span><span class="sxs-lookup"><span data-stu-id="591e1-107">Child elements</span></span>

| <span data-ttu-id="591e1-108">元素</span><span class="sxs-lookup"><span data-stu-id="591e1-108">Element</span></span>                               | <span data-ttu-id="591e1-109">必需</span><span class="sxs-lookup"><span data-stu-id="591e1-109">Required</span></span> | <span data-ttu-id="591e1-110">说明</span><span class="sxs-lookup"><span data-stu-id="591e1-110">Description</span></span>  |
|:--------------------------------------|:--------:|:-------------|
| [<span data-ttu-id="591e1-111">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="591e1-111">ExtensionPoint</span></span>](extensionpoint.md) | <span data-ttu-id="591e1-112">是</span><span class="sxs-lookup"><span data-stu-id="591e1-112">Yes</span></span>      | <span data-ttu-id="591e1-113">定义外接程序公开功能的位置。</span><span class="sxs-lookup"><span data-stu-id="591e1-113">Defines where an add-in exposes functionality.</span></span> |
| [<span data-ttu-id="591e1-114">FunctionFile</span><span class="sxs-lookup"><span data-stu-id="591e1-114">FunctionFile</span></span>](functionfile.md)     | <span data-ttu-id="591e1-115">是</span><span class="sxs-lookup"><span data-stu-id="591e1-115">Yes</span></span>      | <span data-ttu-id="591e1-116">包含 JavaScript 函数的文件的 URL。</span><span class="sxs-lookup"><span data-stu-id="591e1-116">A URL to a file that contains JavaScript functions.</span></span>|
| [<span data-ttu-id="591e1-117">GetStarted</span><span class="sxs-lookup"><span data-stu-id="591e1-117">GetStarted</span></span>](getstarted.md)         | <span data-ttu-id="591e1-118">否</span><span class="sxs-lookup"><span data-stu-id="591e1-118">No</span></span>       | <span data-ttu-id="591e1-119">定义在 Word、Excel 或 PowerPoint 主机中安装外接程序时将显示的标注。</span><span class="sxs-lookup"><span data-stu-id="591e1-119">Defines the callout that appears when installing the add-in in Word, Excel, or PowerPoint hosts.</span></span> |

## <a name="desktopformfactor-example"></a><span data-ttu-id="591e1-120">DesktopFormFactor 示例</span><span class="sxs-lookup"><span data-stu-id="591e1-120">DesktopFormFactor example</span></span>

```xml
...
<Hosts>
  <Host xsi:type="Presentation">
    <DesktopFormFactor>
      <FunctionFile resid="residDesktopFuncUrl" />
      <GetStarted>
        <!-- GetStarted callout -->
      </GetStarted>
      <ExtensionPoint xsi:type="PrimaryCommandSurface">
        <!-- information on this extension point -->
      </ExtensionPoint> 
      <!-- possibly more ExtensionPoint elements -->
    </DesktopFormFactor>
  </Host>
</Hosts>
...
```
