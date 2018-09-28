# <a name="sourcelocation-element"></a><span data-ttu-id="312ee-101">SourceLocation 元素</span><span class="sxs-lookup"><span data-stu-id="312ee-101">SourceLocation element</span></span>

<span data-ttu-id="312ee-p101">指定 Office 外接程序的源文件位置为介于 1 和 2018 个字符之间的 URL。源位置必须是 HTTPS 地址，而非文件路径。</span><span class="sxs-lookup"><span data-stu-id="312ee-p101">Specifies the source file location(s) for your Office Add-in as a URL between 1 and 2018 characters long. The source location must be an HTTPS address, not a file path.</span></span>

<span data-ttu-id="312ee-104">**外接程序类型：** 内容、任务窗格、邮件</span><span class="sxs-lookup"><span data-stu-id="312ee-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="312ee-105">语法</span><span class="sxs-lookup"><span data-stu-id="312ee-105">Syntax</span></span>

```XML
<SourceLocation DefaultValue="string" />
```

## <a name="contained-in"></a><span data-ttu-id="312ee-106">包含在</span><span class="sxs-lookup"><span data-stu-id="312ee-106">Contained in</span></span>

- <span data-ttu-id="312ee-107">[DefaultSettings](defaultsettings.md)（内容和任务窗格外接程序）</span><span class="sxs-lookup"><span data-stu-id="312ee-107">[DefaultSettings](defaultsettings.md) (Content and task pane add-ins)</span></span>
- <span data-ttu-id="312ee-108">[FormSettings](formsettings.md)（邮件外接程序）</span><span class="sxs-lookup"><span data-stu-id="312ee-108">[FormSettings](formsettings.md) (Mail add-ins)</span></span>
- <span data-ttu-id="312ee-109">[ExtensionPoint](extensionpoint.md)（上下文邮件外接程序）</span><span class="sxs-lookup"><span data-stu-id="312ee-109">[ExtensionPoint](extensionpoint.md) (Contextual mail add-ins)</span></span>

## <a name="can-contain"></a><span data-ttu-id="312ee-110">可以包含</span><span class="sxs-lookup"><span data-stu-id="312ee-110">Can contain</span></span>

[<span data-ttu-id="312ee-111">Override</span><span class="sxs-lookup"><span data-stu-id="312ee-111">Override</span></span>](override.md)

## <a name="attributes"></a><span data-ttu-id="312ee-112">属性</span><span class="sxs-lookup"><span data-stu-id="312ee-112">Attributes</span></span>

|<span data-ttu-id="312ee-113">**属性**</span><span class="sxs-lookup"><span data-stu-id="312ee-113">**Attribute**</span></span>|<span data-ttu-id="312ee-114">**类型**</span><span class="sxs-lookup"><span data-stu-id="312ee-114">**Type**</span></span>|<span data-ttu-id="312ee-115">**必需**</span><span class="sxs-lookup"><span data-stu-id="312ee-115">**Required**</span></span>|<span data-ttu-id="312ee-116">**说明**</span><span class="sxs-lookup"><span data-stu-id="312ee-116">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="312ee-117">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="312ee-117">DefaultValue</span></span>|<span data-ttu-id="312ee-118">URL</span><span class="sxs-lookup"><span data-stu-id="312ee-118">URL</span></span>|<span data-ttu-id="312ee-119">必需</span><span class="sxs-lookup"><span data-stu-id="312ee-119">required</span></span>|<span data-ttu-id="312ee-120">指定该设置的默认值，表示为 [DefaultLocale](defaultlocale.md) 元素中指定的区域设置。</span><span class="sxs-lookup"><span data-stu-id="312ee-120">Specifies the default value for this setting for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|
