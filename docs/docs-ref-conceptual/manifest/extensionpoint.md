# <a name="extensionpoint-element"></a><span data-ttu-id="9998c-101">ExtensionPoint 元素</span><span class="sxs-lookup"><span data-stu-id="9998c-101">ExtensionPoint element</span></span>

 <span data-ttu-id="9998c-102">定义 Office UI 中加载项公开功能的位置。</span><span class="sxs-lookup"><span data-stu-id="9998c-102">Defines where an add-in exposes functionality in the Office UI.</span></span> <span data-ttu-id="9998c-103">**ExtensionPoint** 元素是 [AllFormFactors](allformfactors.md)、[DesktopFormFactor](desktopformfactor.md) 或 [MobileFormFactor](mobileformfactor.md) 的子元素。</span><span class="sxs-lookup"><span data-stu-id="9998c-103">The **ExtensionPoint** element is a child element of [AllFormFactors](allformfactors.md), [DesktopFormFactor](desktopformfactor.md) or [MobileFormFactor](mobileformfactor.md).</span></span> 

## <a name="attributes"></a><span data-ttu-id="9998c-104">属性</span><span class="sxs-lookup"><span data-stu-id="9998c-104">Attributes</span></span>

|  <span data-ttu-id="9998c-105">属性</span><span class="sxs-lookup"><span data-stu-id="9998c-105">Attribute</span></span>  |  <span data-ttu-id="9998c-106">必需</span><span class="sxs-lookup"><span data-stu-id="9998c-106">Required</span></span>  |  <span data-ttu-id="9998c-107">说明</span><span class="sxs-lookup"><span data-stu-id="9998c-107">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="9998c-108">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="9998c-108">**xsi:type**</span></span>  |  <span data-ttu-id="9998c-109">是</span><span class="sxs-lookup"><span data-stu-id="9998c-109">Yes</span></span>  | <span data-ttu-id="9998c-110">定义的扩展点类型。</span><span class="sxs-lookup"><span data-stu-id="9998c-110">The type of extension point being defined.</span></span>|

## <a name="extension-points-for-excel-only"></a><span data-ttu-id="9998c-111">仅适用于 Excel 的扩展点</span><span class="sxs-lookup"><span data-stu-id="9998c-111">Extension points for Excel only</span></span>

- <span data-ttu-id="9998c-112">**CustomFunctions** -适用于 Excel 编写 JavaScript 中的自定义函数。</span><span class="sxs-lookup"><span data-stu-id="9998c-112">**CustomFunctions** - A custom function written in JavaScript for Excel.</span></span>

<span data-ttu-id="9998c-113">[此 XML 代码示例](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/customfunctions.xml)演示如何使用**ExtensionPoint**元素的子元素与**CustomFunctions**属性值，用于。</span><span class="sxs-lookup"><span data-stu-id="9998c-113">[This XML code sample](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/customfunctions.xml) shows how to use the **ExtensionPoint** element with the **CustomFunctions** attribute value, and the child elements to be used.</span></span>

## <a name="extension-points-for-word-excel-powerpoint-and-onenote-add-in-commands"></a><span data-ttu-id="9998c-114">适用于 Word、Excel、PowerPoint 和 OneNote 加载项命令的扩展点</span><span class="sxs-lookup"><span data-stu-id="9998c-114">Extension points for Word, Excel, PowerPoint, and OneNote add-in commands</span></span>

- <span data-ttu-id="9998c-115">**PrimaryCommandSurface** - Office 中的功能区。</span><span class="sxs-lookup"><span data-stu-id="9998c-115">**PrimaryCommandSurface** - The ribbon in Office.</span></span>
- <span data-ttu-id="9998c-116">**ContextMenu** - Office UI 中右键单击时出现的快捷菜单。</span><span class="sxs-lookup"><span data-stu-id="9998c-116">**ContextMenu** - The shortcut menu that appears when you right-click in the Office UI.</span></span>

<span data-ttu-id="9998c-117">下面的示例演示如何将  **ExtensionPoint** 元素与 **PrimaryCommandSurface** 和 **ContextMenu** 属性值配合使用，以及应彼此配合使用的子元素。</span><span class="sxs-lookup"><span data-stu-id="9998c-117">The following examples show how to use the  **ExtensionPoint** element with **PrimaryCommandSurface** and **ContextMenu** attribute values, and the child elements that should be used with each.</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="9998c-118">对于包含 ID 属性的元素，请确保您提供一个唯一的 id。</span><span class="sxs-lookup"><span data-stu-id="9998c-118">For elements that contain an ID attribute, make sure you provide a unique ID.</span></span> <span data-ttu-id="9998c-119">建议将你的公司名称用于此 ID。</span><span class="sxs-lookup"><span data-stu-id="9998c-119">We recommend that you use your company's name along with your ID.</span></span> <span data-ttu-id="9998c-120">例如，使用以下格式。</span><span class="sxs-lookup"><span data-stu-id="9998c-120">For example, use the following format.</span></span> <CustomTab id="mycompanyname.mygroupname">

```XML
<ExtensionPoint xsi:type="PrimaryCommandSurface">
          <CustomTab id="Contoso Tab">
          <!-- If you want to use a default tab that comes with Office, remove the above CustomTab element, and then uncomment the following OfficeTab element -->
            <!-- <OfficeTab id="TabData"> -->
            <Label resid="residLabel4" />
            <Group id="Group1Id12">
              <Label resid="residLabel4" />
              <Icon>
                <bt:Image size="16" resid="icon1_32x32" />
                <bt:Image size="32" resid="icon1_32x32" />
                <bt:Image size="80" resid="icon1_32x32" />
              </Icon>
              <Tooltip resid="residToolTip" />
              <Control xsi:type="Button" id="Button1Id1">

                  <!-- information about the control -->
              </Control>
              <!-- other controls, as needed -->
            </Group>
          </CustomTab>
        </ExtensionPoint>

      <ExtensionPoint xsi:type="ContextMenu">
        <OfficeMenu id="ContextMenuCell">
          <Control xsi:type="Menu" id="ContextMenu2">
                  <!-- information about the control -->
          </Control>
          <!-- other controls, as needed -->
        </OfficeMenu>
        </ExtensionPoint>
```

#### <a name="child-elements"></a><span data-ttu-id="9998c-121">子元素</span><span class="sxs-lookup"><span data-stu-id="9998c-121">Child elements</span></span>
 
|<span data-ttu-id="9998c-122">**Element**</span><span class="sxs-lookup"><span data-stu-id="9998c-122">**Element**</span></span>|<span data-ttu-id="9998c-123">**说明**</span><span class="sxs-lookup"><span data-stu-id="9998c-123">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="9998c-124">**CustomTab**</span><span class="sxs-lookup"><span data-stu-id="9998c-124">**CustomTab**</span></span>|<span data-ttu-id="9998c-p103">如果想要（使用 **PrimaryCommandSurface**）向功能区添加自定义选项卡，则为必需项。如果使用 **CustomTab** 元素，则不能使用 **OfficeTab** 元素。**id** 属性是必需的。</span><span class="sxs-lookup"><span data-stu-id="9998c-p103">Required if you want to add a custom tab to the ribbon (using  **PrimaryCommandSurface**). If you use the  **CustomTab** element, you can't use the **OfficeTab** element. The **id** attribute is required.</span></span>|
|<span data-ttu-id="9998c-128">**OfficeTab**</span><span class="sxs-lookup"><span data-stu-id="9998c-128">**OfficeTab**</span></span>|<span data-ttu-id="9998c-p104">如果想要（使用 **PrimaryCommandSurface**）扩展默认 Office 功能区选项卡，则为必需项。如果使用 **OfficeTab** 元素，则不能使用 **CustomTab** 元素。有关详细信息，请参阅 [OfficeTab](officetab.md)。</span><span class="sxs-lookup"><span data-stu-id="9998c-p104">Required if you want to extend a default Office ribbon tab (using **PrimaryCommandSurface**). If you use the  **OfficeTab** element, you can't use the **CustomTab** element. For details, see [OfficeTab](officetab.md).</span></span>|
|<span data-ttu-id="9998c-132">**OfficeMenu**</span><span class="sxs-lookup"><span data-stu-id="9998c-132">**OfficeMenu**</span></span>|<span data-ttu-id="9998c-p105">如果正（使用 **ContextMenu**）将外接程序命令添加到默认上下文菜单中，则为必需项。**id** 属性必须设置为： </span><span class="sxs-lookup"><span data-stu-id="9998c-p105">Required if you're adding add-in commands to a default context menu (using  **ContextMenu**). The  **id** attribute must be set to: </span></span><br/> <span data-ttu-id="9998c-p106">适用于 Excel 或 Word 的 - **ContextMenuText**当用户选定文本，然后右键单击所选定的文本时显示上下文菜单上的项。 </span><span class="sxs-lookup"><span data-stu-id="9998c-p106">- **ContextMenuText** for Excel or Word. Displays the item on the context menu when text is selected and then the user right-clicks on the selected text. </span></span><br/> <span data-ttu-id="9998c-p107">适用于 Excel 的 - **ContextMenuCell**当用户右键单击电子表格中的某个单元格时显示上下文菜单上的项。</span><span class="sxs-lookup"><span data-stu-id="9998c-p107">- **ContextMenuCell** for Excel. Displays the  item on the context menu when the user right-clicks on a cell on the spreadsheet.</span></span>|
|<span data-ttu-id="9998c-139">**Group**</span><span class="sxs-lookup"><span data-stu-id="9998c-139">**Group**</span></span>|<span data-ttu-id="9998c-p108">选项卡上的一组用户界面扩展点。一个组可以有最多六个控件。 **id** 属性是必需项。它是最多使用 125 个字符的字符串。</span><span class="sxs-lookup"><span data-stu-id="9998c-p108">A group of user interface extension points on a tab. A group can have up to six controls. The  **id** attribute is required. It's a string with a maximum of 125 characters.</span></span>|
|<span data-ttu-id="9998c-143">**Label**</span><span class="sxs-lookup"><span data-stu-id="9998c-143">**Label**</span></span>|<span data-ttu-id="9998c-p109">必需。组标签。**resid** 属性必须设置为 **String** 元素的 **id** 属性的值。**String** 元素是 **ShortStrings** 元素的子元素，而 ShortStrings 元素是 **Resources** 元素的子元素。</span><span class="sxs-lookup"><span data-stu-id="9998c-p109">Required. The label of the group. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **ShortStrings** element, which is a child element of the **Resources** element.</span></span>|
|<span data-ttu-id="9998c-148">**Icon**</span><span class="sxs-lookup"><span data-stu-id="9998c-148">**Icon**</span></span>|<span data-ttu-id="9998c-p110">必需。指定将在小型设备上使用或在显示过多按钮的情况下使用的组图标。**resid** 属性必须设置为 **Image** 元素的 **id** 属性的值。**Image** 元素是 **Images** 元素的子元素，而 Images 元素是 **Resources** 元素的子元素。**size** 属性给出图像的大小（以像素为单位）。要求三种图像大小：16、32 和 80。也同样支持五种可选大小：20、24、40、48 和 64。</span><span class="sxs-lookup"><span data-stu-id="9998c-p110">Required. Specifies the group's icon to be used on small form factor devices, or when too many buttons are displayed. The  **resid** attribute must be set to the value of the **id** attribute of an **Image** element. The **Image** element is a child element of the **Images** element, which is a child element of the **Resources** element. The **size** attribute gives the size, in pixels, of the image. Three image sizes are required: 16, 32, and 80. Five optional sizes are also supported: 20, 24, 40, 48, and 64.</span></span>|
|<span data-ttu-id="9998c-156">**Tooltip**</span><span class="sxs-lookup"><span data-stu-id="9998c-156">**Tooltip**</span></span>|<span data-ttu-id="9998c-p111">可选。组的工具提示**resid** 属性必须设置为 **String** 元素的 **id** 属性的值。**String** 元素是 **LongStrings** 元素的子元素，而 LongStrings 元素是 **Resources** 元素的子元素。</span><span class="sxs-lookup"><span data-stu-id="9998c-p111">Optional. The tooltip of the group. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **LongStrings** element, which is a child element of the **Resources** element.</span></span>|
|<span data-ttu-id="9998c-161">**Control**</span><span class="sxs-lookup"><span data-stu-id="9998c-161">**Control**</span></span>|<span data-ttu-id="9998c-162">每个组需要至少一个控件。</span><span class="sxs-lookup"><span data-stu-id="9998c-162">Each group requires at least one control.</span></span> <span data-ttu-id="9998c-163">**Control** 元素可以是一个**按钮**，也可以是一个**菜单**。</span><span class="sxs-lookup"><span data-stu-id="9998c-163">A  **Control** element can be either a **Button** or a **Menu**.</span></span> <span data-ttu-id="9998c-164">使用**菜单**指定按钮控件的下拉列表。</span><span class="sxs-lookup"><span data-stu-id="9998c-164">Use  **Menu** to specify a drop-down list of button controls.</span></span> <span data-ttu-id="9998c-165">目前，仅支持“按钮”和“菜单”。</span><span class="sxs-lookup"><span data-stu-id="9998c-165">Currently, only buttons and menus are supported.</span></span> <span data-ttu-id="9998c-166">请参阅[按钮控件](control.md#button-control)和[菜单控件](control.md#menu-dropdown-button-controls)各节了解详细信息。</span><span class="sxs-lookup"><span data-stu-id="9998c-166">See the [Button controls](control.md#button-control) and [Menu controls](control.md#menu-dropdown-button-controls) sections for more information.</span></span><br/><span data-ttu-id="9998c-167">**注意：** 为了进行故障排除，我们建议将**控件**元素和相关的**资源**子元素添加一次一台。</span><span class="sxs-lookup"><span data-stu-id="9998c-167">**Note:**  To make troubleshooting easier, we recommend that a  **Control** element and the related **Resources** child elements be added one at a time.</span></span>|
|<span data-ttu-id="9998c-168">**Script**</span><span class="sxs-lookup"><span data-stu-id="9998c-168">**Script**</span></span>|<span data-ttu-id="9998c-169">使用自定义函数定义和注册代码链接到 JavaScript 文件。</span><span class="sxs-lookup"><span data-stu-id="9998c-169">Links to the JavaScript file with the custom function definition and registration code.</span></span> <span data-ttu-id="9998c-170">在开发者预览版中不使用此元素。</span><span class="sxs-lookup"><span data-stu-id="9998c-170">This element is not used in the Developer Preview.</span></span> <span data-ttu-id="9998c-171">实际上，HTML 页负责加载所有 JavaScript 文件。</span><span class="sxs-lookup"><span data-stu-id="9998c-171">Instead, the HTML page is responsible for loading all JavaScript files.</span></span>|
|<span data-ttu-id="9998c-172">**Page**</span><span class="sxs-lookup"><span data-stu-id="9998c-172">**Page**</span></span>|<span data-ttu-id="9998c-173">链接到自定义函数的 HTML 页。</span><span class="sxs-lookup"><span data-stu-id="9998c-173">Links to the HTML page for your custom functions.</span></span>|

## <a name="extension-points-for-outlook"></a><span data-ttu-id="9998c-174">Outlook 扩展点</span><span class="sxs-lookup"><span data-stu-id="9998c-174">Extension points for Outlook</span></span>

- [<span data-ttu-id="9998c-175">MessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="9998c-175">MessageReadCommandSurface</span></span>](#messagereadcommandsurface) 
- [<span data-ttu-id="9998c-176">MessageComposeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="9998c-176">MessageComposeCommandSurface</span></span>](#messagecomposecommandsurface) 
- [<span data-ttu-id="9998c-177">AppointmentOrganizerCommandSurface</span><span class="sxs-lookup"><span data-stu-id="9998c-177">AppointmentOrganizerCommandSurface</span></span>](#appointmentorganizercommandsurface) 
- [<span data-ttu-id="9998c-178">AppointmentAttendeeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="9998c-178">AppointmentAttendeeCommandSurface</span></span>](#appointmentattendeecommandsurface)
- <span data-ttu-id="9998c-179">[Module](#module)（仅能在 [DesktopFormFactor](desktopformfactor.md) 中使用。）</span><span class="sxs-lookup"><span data-stu-id="9998c-179">[Module](#module) (Can only be used in the [DesktopFormFactor](desktopformfactor.md).)</span></span>
- [<span data-ttu-id="9998c-180">MobileMessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="9998c-180">MobileMessageReadCommandSurface</span></span>](#mobilemessagereadcommandsurface)
- [<span data-ttu-id="9998c-181">Events</span><span class="sxs-lookup"><span data-stu-id="9998c-181">Events</span></span>](#events)
- [<span data-ttu-id="9998c-182">DetectedEntity</span><span class="sxs-lookup"><span data-stu-id="9998c-182">DetectedEntity</span></span>](#detectedentity)

### <a name="messagereadcommandsurface"></a><span data-ttu-id="9998c-183">MessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="9998c-183">MessageReadCommandSurface</span></span>
<span data-ttu-id="9998c-p114">此扩展点将按钮放置在邮件阅读窗体的命令界面。在 Outlook 桌面，它显示在功能区中。</span><span class="sxs-lookup"><span data-stu-id="9998c-p114">This extension point puts buttons in the command surface for the mail read view. In Outlook desktop, this appears in the ribbon.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="9998c-186">子元素</span><span class="sxs-lookup"><span data-stu-id="9998c-186">Child elements</span></span>

|  <span data-ttu-id="9998c-187">元素</span><span class="sxs-lookup"><span data-stu-id="9998c-187">Element</span></span> |  <span data-ttu-id="9998c-188">说明</span><span class="sxs-lookup"><span data-stu-id="9998c-188">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="9998c-189">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="9998c-189">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="9998c-190">将命令添加到默认功能区选项卡。</span><span class="sxs-lookup"><span data-stu-id="9998c-190">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="9998c-191">CustomTab</span><span class="sxs-lookup"><span data-stu-id="9998c-191">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="9998c-192">将命令添加到自定义功能区选项卡。</span><span class="sxs-lookup"><span data-stu-id="9998c-192">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="9998c-193">OfficeTab 示例</span><span class="sxs-lookup"><span data-stu-id="9998c-193">OfficeTab example</span></span>
```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="9998c-194">CustomTab 示例</span><span class="sxs-lookup"><span data-stu-id="9998c-194">CustomTab example</span></span>
```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="messagecomposecommandsurface"></a><span data-ttu-id="9998c-195">MessageComposeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="9998c-195">MessageComposeCommandSurface</span></span>
<span data-ttu-id="9998c-196">此扩展点将按钮置于使用电子邮件撰写窗体的外接程序的功能区上。</span><span class="sxs-lookup"><span data-stu-id="9998c-196">This extension point puts buttons on the ribbon for add-ins using mail compose form.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="9998c-197">子元素</span><span class="sxs-lookup"><span data-stu-id="9998c-197">Child elements</span></span>

|  <span data-ttu-id="9998c-198">元素</span><span class="sxs-lookup"><span data-stu-id="9998c-198">Element</span></span> |  <span data-ttu-id="9998c-199">说明</span><span class="sxs-lookup"><span data-stu-id="9998c-199">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="9998c-200">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="9998c-200">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="9998c-201">将命令添加到默认功能区选项卡。</span><span class="sxs-lookup"><span data-stu-id="9998c-201">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="9998c-202">CustomTab</span><span class="sxs-lookup"><span data-stu-id="9998c-202">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="9998c-203">将命令添加到自定义功能区选项卡。</span><span class="sxs-lookup"><span data-stu-id="9998c-203">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="9998c-204">OfficeTab 示例</span><span class="sxs-lookup"><span data-stu-id="9998c-204">OfficeTab example</span></span>
```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="9998c-205">CustomTab 示例</span><span class="sxs-lookup"><span data-stu-id="9998c-205">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="appointmentorganizercommandsurface"></a><span data-ttu-id="9998c-206">AppointmentOrganizerCommandSurface</span><span class="sxs-lookup"><span data-stu-id="9998c-206">AppointmentOrganizerCommandSurface</span></span>

<span data-ttu-id="9998c-207">此扩展点将按钮置于向会议的组织者显示的窗体的功能区上。</span><span class="sxs-lookup"><span data-stu-id="9998c-207">This extension point puts buttons on the ribbon for the form that's displayed to the organizer of the meeting.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="9998c-208">子元素</span><span class="sxs-lookup"><span data-stu-id="9998c-208">Child elements</span></span>

|  <span data-ttu-id="9998c-209">元素</span><span class="sxs-lookup"><span data-stu-id="9998c-209">Element</span></span> |  <span data-ttu-id="9998c-210">说明</span><span class="sxs-lookup"><span data-stu-id="9998c-210">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="9998c-211">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="9998c-211">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="9998c-212">将命令添加到默认功能区选项卡。</span><span class="sxs-lookup"><span data-stu-id="9998c-212">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="9998c-213">CustomTab</span><span class="sxs-lookup"><span data-stu-id="9998c-213">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="9998c-214">将命令添加到自定义功能区选项卡。</span><span class="sxs-lookup"><span data-stu-id="9998c-214">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="9998c-215">OfficeTab 示例</span><span class="sxs-lookup"><span data-stu-id="9998c-215">OfficeTab example</span></span>
```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="9998c-216">CustomTab 示例</span><span class="sxs-lookup"><span data-stu-id="9998c-216">CustomTab example</span></span>
```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="appointmentattendeecommandsurface"></a><span data-ttu-id="9998c-217">AppointmentAttendeeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="9998c-217">AppointmentAttendeeCommandSurface</span></span>

<span data-ttu-id="9998c-218">此扩展点将按钮置于向会议与会者显示的窗体的功能区上。</span><span class="sxs-lookup"><span data-stu-id="9998c-218">This extension point puts buttons on the ribbon for the form that's displayed to the attendee of the meeting.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="9998c-219">子元素</span><span class="sxs-lookup"><span data-stu-id="9998c-219">Child elements</span></span>

|  <span data-ttu-id="9998c-220">元素</span><span class="sxs-lookup"><span data-stu-id="9998c-220">Element</span></span> |  <span data-ttu-id="9998c-221">说明</span><span class="sxs-lookup"><span data-stu-id="9998c-221">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="9998c-222">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="9998c-222">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="9998c-223">将命令添加到默认功能区选项卡。</span><span class="sxs-lookup"><span data-stu-id="9998c-223">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="9998c-224">CustomTab</span><span class="sxs-lookup"><span data-stu-id="9998c-224">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="9998c-225">将命令添加到自定义功能区选项卡。</span><span class="sxs-lookup"><span data-stu-id="9998c-225">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="9998c-226">OfficeTab 示例</span><span class="sxs-lookup"><span data-stu-id="9998c-226">OfficeTab example</span></span>
```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="9998c-227">CustomTab 示例</span><span class="sxs-lookup"><span data-stu-id="9998c-227">CustomTab example</span></span>
```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="module"></a><span data-ttu-id="9998c-228">Module</span><span class="sxs-lookup"><span data-stu-id="9998c-228">Module</span></span>

<span data-ttu-id="9998c-229">此扩展点将按钮置于模块扩展的功能区上。</span><span class="sxs-lookup"><span data-stu-id="9998c-229">This extension point puts buttons on the ribbon for the module extension.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="9998c-230">子元素</span><span class="sxs-lookup"><span data-stu-id="9998c-230">Child elements</span></span>

|  <span data-ttu-id="9998c-231">元素</span><span class="sxs-lookup"><span data-stu-id="9998c-231">Element</span></span> |  <span data-ttu-id="9998c-232">说明</span><span class="sxs-lookup"><span data-stu-id="9998c-232">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="9998c-233">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="9998c-233">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="9998c-234">将命令添加到默认功能区选项卡。</span><span class="sxs-lookup"><span data-stu-id="9998c-234">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="9998c-235">CustomTab</span><span class="sxs-lookup"><span data-stu-id="9998c-235">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="9998c-236">将命令添加到自定义功能区选项卡。</span><span class="sxs-lookup"><span data-stu-id="9998c-236">Adds the command(s) to the custom ribbon tab.</span></span>  |

### <a name="mobilemessagereadcommandsurface"></a><span data-ttu-id="9998c-237">MobileMessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="9998c-237">MobileMessageReadCommandSurface</span></span>
<span data-ttu-id="9998c-238">此扩展点将按钮置于移动外形规格中的邮件阅读视图的命令界面中。</span><span class="sxs-lookup"><span data-stu-id="9998c-238">This extension point puts buttons in the command surface for the mail read view in the mobile form factor.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="9998c-239">子元素</span><span class="sxs-lookup"><span data-stu-id="9998c-239">Child elements</span></span>

|  <span data-ttu-id="9998c-240">元素</span><span class="sxs-lookup"><span data-stu-id="9998c-240">Element</span></span> |  <span data-ttu-id="9998c-241">说明</span><span class="sxs-lookup"><span data-stu-id="9998c-241">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="9998c-242">Group</span><span class="sxs-lookup"><span data-stu-id="9998c-242">Group</span></span>](group.md) |  <span data-ttu-id="9998c-243">将按钮组添加到命令界面。</span><span class="sxs-lookup"><span data-stu-id="9998c-243">Adds a group of buttons to the command surface.</span></span>  |

<span data-ttu-id="9998c-244">此种类型的 **ExtensionPoint** 元素仅能具有一个子元素，即 **Group** 元素。</span><span class="sxs-lookup"><span data-stu-id="9998c-244">**ExtensionPoint** elements of this type can only have one child element: a **Group** element.</span></span>

<span data-ttu-id="9998c-245">此扩展点中包含的 **Control** 元素必须将 **xsi:type** 属性设置为 `MobileButton`。</span><span class="sxs-lookup"><span data-stu-id="9998c-245">**Control** elements contained in this extension point must have the **xsi:type** attribute set to `MobileButton`.</span></span>

#### <a name="example"></a><span data-ttu-id="9998c-246">示例</span><span class="sxs-lookup"><span data-stu-id="9998c-246">Example</span></span>
```xml
<ExtensionPoint xsi:type="MobileMessageReadCommandSurface">
  <Group id="mobileGroupID">
    <Label resid="residAppName"/>
      <Control id="mobileButton1" xsi:type="MobileButton">
        <!-- Control definition -->
      </Control>
  </Group>
</ExtensionPoint>
```

### <a name="events"></a><span data-ttu-id="9998c-247">事件</span><span class="sxs-lookup"><span data-stu-id="9998c-247">Events</span></span>

<span data-ttu-id="9998c-248">此扩展点添加了指定事件的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="9998c-248">This extension point adds an event handler for a specified event.</span></span>

> [!NOTE]
> <span data-ttu-id="9998c-249">此元素类型只支持 Office 365 中 web 上的 outlook 中。</span><span class="sxs-lookup"><span data-stu-id="9998c-249">This element type is only supported by Outlook on the web in Office 365.</span></span>

| <span data-ttu-id="9998c-250">元素</span><span class="sxs-lookup"><span data-stu-id="9998c-250">Element</span></span> | <span data-ttu-id="9998c-251">说明</span><span class="sxs-lookup"><span data-stu-id="9998c-251">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="9998c-252">Event</span><span class="sxs-lookup"><span data-stu-id="9998c-252">Event</span></span>](event.md) |  <span data-ttu-id="9998c-253">指定事件和事件处理程序函数。</span><span class="sxs-lookup"><span data-stu-id="9998c-253">Specifies the event and event handler function.</span></span>  |

#### <a name="itemsend-event-example"></a><span data-ttu-id="9998c-254">ItemSend 事件示例</span><span class="sxs-lookup"><span data-stu-id="9998c-254">ItemSend event example</span></span>

```xml
<ExtensionPoint xsi:type="Events"> 
  <Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="itemSendHandler" /> 
</ExtensionPoint> 
```

### <a name="detectedentity"></a><span data-ttu-id="9998c-255">DetectedEntity</span><span class="sxs-lookup"><span data-stu-id="9998c-255">DetectedEntity</span></span>

<span data-ttu-id="9998c-256">此扩展点在指定实体类型上添加上下文外接程序激活。</span><span class="sxs-lookup"><span data-stu-id="9998c-256">This extension point adds a contextual add-in activation on a specified entity type.</span></span>

<span data-ttu-id="9998c-257">包含的 [VersionOverrides](versionoverrides.md) 元素的 `VersionOverridesV1_1` 的属性值必须为 `xsi:type`。</span><span class="sxs-lookup"><span data-stu-id="9998c-257">The containing [VersionOverrides](versionoverrides.md) element must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span>

> [!NOTE]
> <span data-ttu-id="9998c-258">此元素类型只支持 Office 365 中 web 上的 outlook 中。</span><span class="sxs-lookup"><span data-stu-id="9998c-258">This element type is only supported by Outlook on the web in Office 365.</span></span>

|  <span data-ttu-id="9998c-259">元素</span><span class="sxs-lookup"><span data-stu-id="9998c-259">Element</span></span> |  <span data-ttu-id="9998c-260">说明</span><span class="sxs-lookup"><span data-stu-id="9998c-260">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="9998c-261">Label</span><span class="sxs-lookup"><span data-stu-id="9998c-261">Label</span></span>](#label) |  <span data-ttu-id="9998c-262">在上下文窗口中指定外接程序的标签。</span><span class="sxs-lookup"><span data-stu-id="9998c-262">Specifies the label for the add-in in the contextual window.</span></span>  |
|  [<span data-ttu-id="9998c-263">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="9998c-263">SourceLocation</span></span>](sourcelocation.md) |  <span data-ttu-id="9998c-264">指定上下文窗口的 URL。</span><span class="sxs-lookup"><span data-stu-id="9998c-264">Specifies the URL for the contextual window.</span></span>  |
|  [<span data-ttu-id="9998c-265">Rule</span><span class="sxs-lookup"><span data-stu-id="9998c-265">Rule</span></span>](rule.md) |  <span data-ttu-id="9998c-266">指定确定外接程序激活时间的一个或多个规则。</span><span class="sxs-lookup"><span data-stu-id="9998c-266">Specifies the rule or rules that determine when an add-in activates.</span></span>  |

#### <a name="label"></a><span data-ttu-id="9998c-267">标签</span><span class="sxs-lookup"><span data-stu-id="9998c-267">Label</span></span>

<span data-ttu-id="9998c-p115">必需。组的标签。 **resid** 属性必须设置为 **ShortStrings** 元素（位于 **Resources** 元素）中 **String** 元素的 [id](resources.md) 属性的值。</span><span class="sxs-lookup"><span data-stu-id="9998c-p115">Required. The label of the group. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

#### <a name="highlight-requirements"></a><span data-ttu-id="9998c-271">突出显示要求</span><span class="sxs-lookup"><span data-stu-id="9998c-271">Highlight requirements</span></span>

<span data-ttu-id="9998c-p116">用户可以激活上下文外接程序的唯一方法是与突出显示实体进行交互。开发人员可以使用 `ItemHasKnownEntity` 和`ItemHasRegularExpressionMatch` 规则类型的 `Rule` 元素的 `Highlight` 属性来控制突出显示哪些实体。</span><span class="sxs-lookup"><span data-stu-id="9998c-p116">The only way a user can activate a contextual add-in is to interact with a highlighted entity. Developers can control which entities are highlighted by using the `Highlight` attribute of the `Rule` element for `ItemHasKnownEntity` and `ItemHasRegularExpressionMatch` rule types.</span></span>

<span data-ttu-id="9998c-p117">但是，存在一些需要注意的限制。存在这些限制是为了确保在适用的邮件或约会中始终存在一个突出显示实体，以便为用户提供一种激活外接程序的方法。</span><span class="sxs-lookup"><span data-stu-id="9998c-p117">However, there are some limitations to be aware of. These limitations are in place to ensure that there will always be a highlighted entity in applicable messages or appointments to give the user a way to activate the add-in.</span></span>

- <span data-ttu-id="9998c-276">无法突出显示 `EmailAddress` 和 `Url` 实体类型，因此不能用于激活外接程序。</span><span class="sxs-lookup"><span data-stu-id="9998c-276">The `EmailAddress` and `Url` entity types cannot be highlighted, and therefore cannot be used to activate an add-in.</span></span>
- <span data-ttu-id="9998c-277">如果使用单个规则，`Highlight` 必须设置为 `all`。</span><span class="sxs-lookup"><span data-stu-id="9998c-277">If using a single rule, `Highlight` MUST be set to `all`.</span></span>
- <span data-ttu-id="9998c-278">如果使用具有 `Mode="AND"` 的 `RuleCollection` 规则类型来组合多个规则，则至少其中有一个规则必须将 `Highlight` 设置为 `all`。</span><span class="sxs-lookup"><span data-stu-id="9998c-278">If using a `RuleCollection` rule type with `Mode="AND"` to combine multiple rules, at least one of the rules MUST have `Highlight` set to `all`.</span></span>
- <span data-ttu-id="9998c-279">如果使用具有 `Mode="OR"` 的 `RuleCollection` 规则类型来组合多个规则，则所有规则必须将 `Highlight` 设置为 `all`。</span><span class="sxs-lookup"><span data-stu-id="9998c-279">If using a `RuleCollection` rule type with `Mode="OR"` to combine multiple rules, all of the rules MUST have `Highlight` set to `all`.</span></span>

#### <a name="detectedentity-event-example"></a><span data-ttu-id="9998c-280">DetectedEntity 事件示例</span><span class="sxs-lookup"><span data-stu-id="9998c-280">DetectedEntity event example</span></span>

```xml
<ExtensionPoint xsi:type="DetectedEntity">
  <Label resid="residLabelName"/>
  <SourceLocation resid="residDetectedEntityURL" />
  <Rule xsi:type="RuleCollection" Mode="And">
    <Rule xsi:type="ItemIs" ItemType="Message" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="MeetingSuggestion" Highlight="all" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" Highlight="none" />
  </Rule>
</ExtensionPoint> 
```