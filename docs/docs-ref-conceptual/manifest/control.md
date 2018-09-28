# <a name="control-element"></a><span data-ttu-id="87574-101">Control 元素</span><span class="sxs-lookup"><span data-stu-id="87574-101">Control element</span></span>

<span data-ttu-id="87574-p101">定义执行操作或启动任务窗格的 JavaScript 函数。**Control** 元素可以是按钮选项，也可以是菜单选项。[Group](group.md) 元素中至少需包括一个 **Control**。</span><span class="sxs-lookup"><span data-stu-id="87574-p101">Defines a JavaScript function that executes an action or launches a task pane. A **Control** element can be either a button or a menu option. At least one **Control** must be included in a [Group](group.md) element.</span></span>

## <a name="attributes"></a><span data-ttu-id="87574-105">属性</span><span class="sxs-lookup"><span data-stu-id="87574-105">Attributes</span></span>

|  <span data-ttu-id="87574-106">属性</span><span class="sxs-lookup"><span data-stu-id="87574-106">Attribute</span></span>  |  <span data-ttu-id="87574-107">必需</span><span class="sxs-lookup"><span data-stu-id="87574-107">Required</span></span>  |  <span data-ttu-id="87574-108">说明</span><span class="sxs-lookup"><span data-stu-id="87574-108">Description</span></span>  |
|:-----|:-----|:-----|
|<span data-ttu-id="87574-109">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="87574-109">**xsi:type**</span></span>|<span data-ttu-id="87574-110">是</span><span class="sxs-lookup"><span data-stu-id="87574-110">Yes</span></span>|<span data-ttu-id="87574-p102">正在定义的控件类型。可以是 `Button`、`Menu` 或 `MobileButton`。</span><span class="sxs-lookup"><span data-stu-id="87574-p102">The type of control being defined. Can be either `Button`, `Menu`, or `MobileButton`.</span></span> |
|<span data-ttu-id="87574-113">**id**</span><span class="sxs-lookup"><span data-stu-id="87574-113">**id**</span></span>|<span data-ttu-id="87574-114">否</span><span class="sxs-lookup"><span data-stu-id="87574-114">No</span></span>|<span data-ttu-id="87574-p103">控件元素的 ID。最多可包含 125 个字符。</span><span class="sxs-lookup"><span data-stu-id="87574-p103">The ID of the control element. Can be a maximum of 125 characters.</span></span>|

> [!NOTE]
> <span data-ttu-id="87574-117">`MobileButton`的**xsi: type** VersionOverrides 架构 1.1 中定义的值。</span><span class="sxs-lookup"><span data-stu-id="87574-117">The `MobileButton` value for **xsi:type** is defined in VersionOverrides schema 1.1.</span></span> <span data-ttu-id="87574-118">它仅适用于[MobileFormFactor](mobileformfactor.md)元素中包含的**控件**元素。</span><span class="sxs-lookup"><span data-stu-id="87574-118">It only applies to the **Control** elements contained within a [MobileFormFactor](mobileformfactor.md) element.</span></span>

## <a name="button-control"></a><span data-ttu-id="87574-119">按钮控件</span><span class="sxs-lookup"><span data-stu-id="87574-119">Button control</span></span>

<span data-ttu-id="87574-p105">当用户选择某个按钮时，将执行一个操作。它可以执行函数或显示任务窗格。每个按钮控件必须具有对清单唯一的 `id`。</span><span class="sxs-lookup"><span data-stu-id="87574-p105">A button performs a single action when the user selects it. It can either execute a function or show a task pane. Each button control must have an `id` unique to the manifest.</span></span> 

### <a name="child-elements"></a><span data-ttu-id="87574-123">子元素</span><span class="sxs-lookup"><span data-stu-id="87574-123">Child elements</span></span>
|  <span data-ttu-id="87574-124">元素</span><span class="sxs-lookup"><span data-stu-id="87574-124">Element</span></span> |  <span data-ttu-id="87574-125">必需</span><span class="sxs-lookup"><span data-stu-id="87574-125">Required</span></span>  |  <span data-ttu-id="87574-126">说明</span><span class="sxs-lookup"><span data-stu-id="87574-126">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="87574-127">**标签**</span><span class="sxs-lookup"><span data-stu-id="87574-127">**Label**</span></span>     | <span data-ttu-id="87574-128">是</span><span class="sxs-lookup"><span data-stu-id="87574-128">Yes</span></span> |  <span data-ttu-id="87574-p106">按钮文本。**resid** 属性必须设置为 **ShortStrings** 元素（位于 [Resources](resources.md) 元素）中 **String** 元素的 **id** 属性的值。</span><span class="sxs-lookup"><span data-stu-id="87574-p106">The text for the button. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md)  element.</span></span>        |
|  <span data-ttu-id="87574-131">**ToolTip**</span><span class="sxs-lookup"><span data-stu-id="87574-131">**ToolTip**</span></span>  |<span data-ttu-id="87574-132">否</span><span class="sxs-lookup"><span data-stu-id="87574-132">No</span></span>|<span data-ttu-id="87574-p107">按钮的工具提示。**resid** 属性必须设置为 **String** 元素的 **id** 属性的值。**String** 元素是 **LongStrings** 元素的子元素，而 LongStrings 元素是 [Resources](resources.md) 元素的子元素。</span><span class="sxs-lookup"><span data-stu-id="87574-p107">The tooltip for the button. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child of the **LongStrings** element, which is a child of the [Resources](resources.md) element.</span></span>|        
|  [<span data-ttu-id="87574-136">Supertip</span><span class="sxs-lookup"><span data-stu-id="87574-136">Supertip</span></span>](supertip.md)  | <span data-ttu-id="87574-137">是</span><span class="sxs-lookup"><span data-stu-id="87574-137">Yes</span></span> |  <span data-ttu-id="87574-138">按钮的 supertip。</span><span class="sxs-lookup"><span data-stu-id="87574-138">The supertip for the button.</span></span>    |
|  [<span data-ttu-id="87574-139">图标</span><span class="sxs-lookup"><span data-stu-id="87574-139">Icon</span></span>](icon.md)      | <span data-ttu-id="87574-140">是</span><span class="sxs-lookup"><span data-stu-id="87574-140">Yes</span></span> |  <span data-ttu-id="87574-141">按钮的图像。</span><span class="sxs-lookup"><span data-stu-id="87574-141">An image for the button.</span></span>         |
|  [<span data-ttu-id="87574-142">Action</span><span class="sxs-lookup"><span data-stu-id="87574-142">Action</span></span>](action.md)    | <span data-ttu-id="87574-143">是</span><span class="sxs-lookup"><span data-stu-id="87574-143">Yes</span></span> |  <span data-ttu-id="87574-144">指定要执行的操作。</span><span class="sxs-lookup"><span data-stu-id="87574-144">Specifies the action to perform.</span></span>  |

### <a name="executefunction-button-example"></a><span data-ttu-id="87574-145">ExecuteFunction 按钮示例</span><span class="sxs-lookup"><span data-stu-id="87574-145">ExecuteFunction button example</span></span>

```xml
<Control xsi:type="Button" id="msgReadFunctionButton">
  <Label resid="funcReadButtonLabel" />
  <Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="blue-icon-16" />
    <bt:Image size="32" resid="blue-icon-32" />
    <bt:Image size="80" resid="blue-icon-80" />
  </Icon>
  <Action xsi:type="ExecuteFunction">
    <FunctionName>getSubject</FunctionName>
  </Action>
</Control>
```

### <a name="showtaskpane-button-example"></a><span data-ttu-id="87574-146">ShowTaskpane 按钮示例</span><span class="sxs-lookup"><span data-stu-id="87574-146">ShowTaskpane button example</span></span>

```xml
<Control xsi:type="Button" id="msgReadOpenPaneButton">
  <Label resid="paneReadButtonLabel" />
  <Supertip>
    <Title resid="paneReadSuperTipTitle" />
    <Description resid="paneReadSuperTipDescription" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="green-icon-16" />
    <bt:Image size="32" resid="green-icon-32" />
    <bt:Image size="80" resid="green-icon-80" />
  </Icon>
  <Action xsi:type="ShowTaskpane">
    <SourceLocation resid="readTaskPaneUrl" />
  </Action>
</Control>
```

## <a name="menu-dropdown-button-controls"></a><span data-ttu-id="87574-147">菜单（下拉）控件</span><span class="sxs-lookup"><span data-stu-id="87574-147">Menu (dropdown button) controls</span></span>

<span data-ttu-id="87574-p108">菜单定义选项的静态列表。每个菜单项将执行函数或显示任务窗格。不支持子菜单。</span><span class="sxs-lookup"><span data-stu-id="87574-p108">A menu defines a static list of options. Each menu item either executes a function or shows a task pane. Submenus are not supported.</span></span> 

<span data-ttu-id="87574-151">使用 **PrimaryCommandSurface** 或 **ContextMenu** [扩展点](extensionpoint.md)时，菜单控件定义：</span><span class="sxs-lookup"><span data-stu-id="87574-151">When used with a **PrimaryCommandSurface** or **ContextMenu** [extension point](extensionpoint.md), the menu control defines:</span></span>

- <span data-ttu-id="87574-152">根级别菜单项。</span><span class="sxs-lookup"><span data-stu-id="87574-152">A root-level menu item.</span></span>

- <span data-ttu-id="87574-153">子菜单项的列表。</span><span class="sxs-lookup"><span data-stu-id="87574-153">A list of submenu items.</span></span>

<span data-ttu-id="87574-p109">当与  **PrimaryCommandSurface** 一起使用时，根菜单项将显示为功能区上的按钮。选择该按钮后，子菜单将显示为下拉列表。与 **ContextMenu** 一起使用时，具有子菜单的菜单项将被插入到上下文菜单上。在这两种情况下，单个子菜单项可以执行 JavaScript 函数，也可显示任务窗格。这一次仅支持子菜单的一个级别。</span><span class="sxs-lookup"><span data-stu-id="87574-p109">When used with **PrimaryCommandSurface**, the root menu item displays as a button on the ribbon. When the button is selected, the submenu displays as a drop-down list. When used with  **ContextMenu**, a menu item with a submenu is inserted on the context menu. In both cases, individual submenu items can either execute a JavaScript function or show a task pane. Only one level of submenus is supported at this time.</span></span>

<span data-ttu-id="87574-p110">下面的示例演示如何定义具有两个子菜单项的菜单项。第一个子菜单项显示任务窗格，而第二个子菜单项运行 JavaScript 函数。</span><span class="sxs-lookup"><span data-stu-id="87574-p110">The following example shows how to define a menu item with two submenu items. The first submenu item shows a task pane, and the second submenu item runs a JavaScript function.</span></span>

```xml
<Control xsi:type="Menu" id="TestMenu2">
  <Label resid="residLabel3" />
  <Tooltip resid="residToolTip" />
  <Supertip>
    <Title resid="residLabel" />
    <Description resid="residToolTip" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="icon1_32x32" />
    <bt:Image size="32" resid="icon1_32x32" />
    <bt:Image size="80" resid="icon1_32x32" />
  </Icon>
  <Items>
    <Item id="showGallery2">
      <Label resid="residLabel3"/>
      <Supertip>
        <Title resid="residLabel" />
        <Description resid="residToolTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="icon1_32x32" />
        <bt:Image size="32" resid="icon1_32x32" />
        <bt:Image size="80" resid="icon1_32x32" />
      </Icon>
      <Action xsi:type="ShowTaskpane">
        <TaskpaneId>MyTaskPaneID1</TaskpaneId>
        <SourceLocation resid="residUnitConverterUrl" />
      </Action>
    </Item>
    <Item id="showGallery3">
      <Label resid="residLabel5"/>
      <Supertip>
        <Title resid="residLabel" />
        <Description resid="residToolTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="icon4_32x32" />
        <bt:Image size="32" resid="icon4_32x32" />
        <bt:Image size="80" resid="icon4_32x32" />
      </Icon>
      <Action xsi:type="ExecuteFunction">
        <FunctionName>getButton</FunctionName>
      </Action>
    </Item>
  </Items>
</Control>

```

### <a name="child-elements"></a><span data-ttu-id="87574-161">子元素</span><span class="sxs-lookup"><span data-stu-id="87574-161">Child elements</span></span>

|  <span data-ttu-id="87574-162">元素</span><span class="sxs-lookup"><span data-stu-id="87574-162">Element</span></span> |  <span data-ttu-id="87574-163">必需</span><span class="sxs-lookup"><span data-stu-id="87574-163">Required</span></span>  |  <span data-ttu-id="87574-164">说明</span><span class="sxs-lookup"><span data-stu-id="87574-164">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="87574-165">**标签**</span><span class="sxs-lookup"><span data-stu-id="87574-165">**Label**</span></span>     | <span data-ttu-id="87574-166">是</span><span class="sxs-lookup"><span data-stu-id="87574-166">Yes</span></span> |  <span data-ttu-id="87574-p111">按钮文本。**resid** 属性必须设置为 **ShortStrings** 元素（位于 [Resources](resources.md) 元素）中 **String** 元素的 **id** 属性的值。</span><span class="sxs-lookup"><span data-stu-id="87574-p111">The text for the button. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>      |
|  <span data-ttu-id="87574-169">**ToolTip**</span><span class="sxs-lookup"><span data-stu-id="87574-169">**ToolTip**</span></span>  |<span data-ttu-id="87574-170">否</span><span class="sxs-lookup"><span data-stu-id="87574-170">No</span></span>|<span data-ttu-id="87574-p112">按钮的工具提示。**resid** 属性必须设置为 **String** 元素的 **id** 属性的值。**String** 元素是 **LongStrings** 元素的子元素，而 LongStrings 元素是 [Resources](resources.md) 元素的子元素。</span><span class="sxs-lookup"><span data-stu-id="87574-p112">The tooltip for the button. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child of the **LongStrings** element, which is a child of the [Resources](resources.md) element.</span></span>|        
|  [<span data-ttu-id="87574-174">Supertip</span><span class="sxs-lookup"><span data-stu-id="87574-174">Supertip</span></span>](supertip.md)  | <span data-ttu-id="87574-175">是</span><span class="sxs-lookup"><span data-stu-id="87574-175">Yes</span></span> |  <span data-ttu-id="87574-176">此按钮的 supertip。</span><span class="sxs-lookup"><span data-stu-id="87574-176">The supertip for this button.</span></span>    |
|  [<span data-ttu-id="87574-177">图标</span><span class="sxs-lookup"><span data-stu-id="87574-177">Icon</span></span>](icon.md)      | <span data-ttu-id="87574-178">是</span><span class="sxs-lookup"><span data-stu-id="87574-178">Yes</span></span> |  <span data-ttu-id="87574-179">按钮的图像。</span><span class="sxs-lookup"><span data-stu-id="87574-179">An image for the button.</span></span>         |
|  <span data-ttu-id="87574-180">**Items**</span><span class="sxs-lookup"><span data-stu-id="87574-180">**Items**</span></span>     | <span data-ttu-id="87574-181">是</span><span class="sxs-lookup"><span data-stu-id="87574-181">Yes</span></span> |  <span data-ttu-id="87574-p113">菜单中显示的按钮的集合。包含每个子菜单项的 **Item** 元素。每个 **Item** 元素均包含 [按钮控件](#button-control) 的子元素。</span><span class="sxs-lookup"><span data-stu-id="87574-p113">A collection of Buttons to display within the menu. Contains the  **Item** elements for each submenu item. Each **Item** element contains the  child elements of the [Button control](#button-control).</span></span>|

### <a name="menu-control-examples"></a><span data-ttu-id="87574-185">菜单控件示例</span><span class="sxs-lookup"><span data-stu-id="87574-185">Menu control examples</span></span>

```xml
<Control xsi:type="Menu" id="TestMenu2">
  <Label resid="residLabel3" />
  <Tooltip resid="residToolTip" />
  <Supertip>
    <Title resid="residLabel" />
    <Description resid="residToolTip" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="icon1_32x32" />
    <bt:Image size="32" resid="icon1_32x32" />
    <bt:Image size="80" resid="icon1_32x32" />
  </Icon>
  <Items>
    <Item id="showGallery2">
      <Label resid="residLabel3"/>
      <Supertip>
        <Title resid="residLabel" />
        <Description resid="residToolTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="icon1_32x32" />
        <bt:Image size="32" resid="icon1_32x32" />
        <bt:Image size="80" resid="icon1_32x32" />
      </Icon>
      <Action xsi:type="ShowTaskpane">
        <TaskpaneId>MyTaskPaneID1</TaskpaneId>
        <SourceLocation resid="residUnitConverterUrl" />
      </Action>
    </Item>
    <Item id="showGallery3">
      <Label resid="residLabel5"/>
      <Supertip>
        <Title resid="residLabel" />
        <Description resid="residToolTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="icon4_32x32" />
        <bt:Image size="32" resid="icon4_32x32" />
        <bt:Image size="80" resid="icon4_32x32" />
      </Icon>
      <Action xsi:type="ExecuteFunction">
        <FunctionName>getButton</FunctionName>
      </Action>
    </Item>
  </Items>
</Control>

```

```xml
<Control xsi:type="Menu" id="msgReadMenuButton">
  <Label resid="menuReadButtonLabel" />
  <Supertip>
    <Title resid="menuReadSuperTipTitle" />
    <Description resid="menuReadSuperTipDescription" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="red-icon-16" />
    <bt:Image size="32" resid="red-icon-32" />
    <bt:Image size="80" resid="red-icon-80" />
  </Icon>
  <Items>
    <Item id="msgReadMenuItem1">
      <Label resid="menuItem1ReadLabel" />
      <Supertip>
        <Title resid="menuItem1ReadLabel" />
        <Description resid="menuItem1ReadTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="red-icon-16" />
        <bt:Image size="32" resid="red-icon-32" />
        <bt:Image size="80" resid="red-icon-80" />
      </Icon>
      <Action xsi:type="ExecuteFunction">
        <FunctionName>getItemClass</FunctionName>
      </Action>
    </Item>
  </Items>
</Control>
```

## <a name="mobilebutton-control"></a><span data-ttu-id="87574-186">MobileButton 控件</span><span class="sxs-lookup"><span data-stu-id="87574-186">MobileButton control</span></span>

<span data-ttu-id="87574-p114">当用户选择某个移动按钮时，将执行一个操作。它可以执行函数或显示任务窗格。 每个移动按钮控件必须具有对清单唯一的 `id`。</span><span class="sxs-lookup"><span data-stu-id="87574-p114">A mobile button performs a single action when the user selects it. It can either execute a function or show a task pane. Each mobile button control must have an `id` unique to the manifest.</span></span>

<span data-ttu-id="87574-p115">在 VersionOverrides 架构 1.1 中定义了 **xsi:type** 的 `MobileButton` 值。包含  [VersionOverrides](versionoverrides.md) 元素的 `VersionOverridesV1_1` 属性值必须为 `xsi:type`。</span><span class="sxs-lookup"><span data-stu-id="87574-p115">The `MobileButton` value for **xsi:type** is defined in VersionOverrides schema 1.1. The containing [VersionOverrides](versionoverrides.md) element must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span>

### <a name="child-elements"></a><span data-ttu-id="87574-192">子元素</span><span class="sxs-lookup"><span data-stu-id="87574-192">Child elements</span></span>
|  <span data-ttu-id="87574-193">元素</span><span class="sxs-lookup"><span data-stu-id="87574-193">Element</span></span> |  <span data-ttu-id="87574-194">必需</span><span class="sxs-lookup"><span data-stu-id="87574-194">Required</span></span>  |  <span data-ttu-id="87574-195">说明</span><span class="sxs-lookup"><span data-stu-id="87574-195">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="87574-196">**标签**</span><span class="sxs-lookup"><span data-stu-id="87574-196">**Label**</span></span>     | <span data-ttu-id="87574-197">是</span><span class="sxs-lookup"><span data-stu-id="87574-197">Yes</span></span> |  <span data-ttu-id="87574-p116">按钮文本。**resid** 属性必须设置为 **ShortStrings** 元素（位于 [Resources](resources.md) 元素）中 **String** 元素的 **id** 属性的值。</span><span class="sxs-lookup"><span data-stu-id="87574-p116">The text for the button. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md)  element.</span></span>        |
|  [<span data-ttu-id="87574-200">图标</span><span class="sxs-lookup"><span data-stu-id="87574-200">Icon</span></span>](icon.md)      | <span data-ttu-id="87574-201">是</span><span class="sxs-lookup"><span data-stu-id="87574-201">Yes</span></span> |  <span data-ttu-id="87574-202">按钮的图像。</span><span class="sxs-lookup"><span data-stu-id="87574-202">An image for the button.</span></span>         |
|  [<span data-ttu-id="87574-203">Action</span><span class="sxs-lookup"><span data-stu-id="87574-203">Action</span></span>](action.md)    | <span data-ttu-id="87574-204">是</span><span class="sxs-lookup"><span data-stu-id="87574-204">Yes</span></span> |  <span data-ttu-id="87574-205">指定要执行的操作。</span><span class="sxs-lookup"><span data-stu-id="87574-205">Specifies the action to perform.</span></span>  |

### <a name="executefunction-mobile-button-example"></a><span data-ttu-id="87574-206">ExecuteFunction 移动按钮示例</span><span class="sxs-lookup"><span data-stu-id="87574-206">ExecuteFunction mobile button example</span></span>

```xml
<Control xsi:type="MobileButton" id="msgReadFunctionButton">
  <Label resid="funcReadButtonLabel" />
  <Icon>
    <bt:Image resid="blue-icon-16-1" size="25" scale="1" />
    <bt:Image resid="blue-icon-16-2" size="25" scale="2" />
    <bt:Image resid="blue-icon-16-3" size="25" scale="3" />
    <bt:Image resid="blue-icon-32-1" size="32" scale="1" />
    <bt:Image resid="blue-icon-32-2" size="32" scale="2" />
    <bt:Image resid="blue-icon-32-3" size="32" scale="3" />
    <bt:Image resid="blue-icon-80-1" size="48" scale="1" />
    <bt:Image resid="blue-icon-80-2" size="48" scale="2" />
    <bt:Image resid="blue-icon-80-3" size="48" scale="3" />
  </Icon>
  <Action xsi:type="ExecuteFunction">
    <FunctionName>getSubject</FunctionName>
  </Action>
</Control>
```

### <a name="showtaskpane-mobile-button-example"></a><span data-ttu-id="87574-207">ShowTaskpane 移动按钮示例</span><span class="sxs-lookup"><span data-stu-id="87574-207">ShowTaskpane mobile button example</span></span>

```xml
<Control xsi:type="MobileButton" id="msgReadOpenPaneButton">
  <Label resid="paneReadButtonLabel" />
  <Icon>
    <bt:Image resid="blue-icon-16-1" size="25" scale="1" />
    <bt:Image resid="blue-icon-16-2" size="25" scale="2" />
    <bt:Image resid="blue-icon-16-3" size="25" scale="3" />
    <bt:Image resid="blue-icon-32-1" size="32" scale="1" />
    <bt:Image resid="blue-icon-32-2" size="32" scale="2" />
    <bt:Image resid="blue-icon-32-3" size="32" scale="3" />
    <bt:Image resid="blue-icon-80-1" size="48" scale="1" />
    <bt:Image resid="blue-icon-80-2" size="48" scale="2" />
    <bt:Image resid="blue-icon-80-3" size="48" scale="3" />
  </Icon>
  <Action xsi:type="ShowTaskpane">
    <SourceLocation resid="readTaskPaneUrl" />
  </Action>
</Control>
```