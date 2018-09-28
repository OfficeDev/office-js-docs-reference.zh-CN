# <a name="method-element"></a><span data-ttu-id="1cbaa-101">Method 元素</span><span class="sxs-lookup"><span data-stu-id="1cbaa-101">Method element</span></span>

<span data-ttu-id="1cbaa-102">指定来自适用于 Office 的 JavaScript API 的单个方法，Office 外接程序需要该方法才能激活。</span><span class="sxs-lookup"><span data-stu-id="1cbaa-102">Specifies an individual method from the JavaScript API for Office that your Office Add-in requires in order to activate.</span></span>

<span data-ttu-id="1cbaa-103">**外接程序类型：** 内容、任务窗格</span><span class="sxs-lookup"><span data-stu-id="1cbaa-103">**Add-in type:** Content, Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="1cbaa-104">语法</span><span class="sxs-lookup"><span data-stu-id="1cbaa-104">Syntax</span></span>

```XML
<Method Name="string"/>
```

## <a name="contained-in"></a><span data-ttu-id="1cbaa-105">包含在</span><span class="sxs-lookup"><span data-stu-id="1cbaa-105">Contained in</span></span>

[<span data-ttu-id="1cbaa-106">方法</span><span class="sxs-lookup"><span data-stu-id="1cbaa-106">Methods</span></span>](methods.md)

## <a name="attributes"></a><span data-ttu-id="1cbaa-107">Attributes</span><span class="sxs-lookup"><span data-stu-id="1cbaa-107">Attributes</span></span>

|<span data-ttu-id="1cbaa-108">**属性**</span><span class="sxs-lookup"><span data-stu-id="1cbaa-108">**Attribute**</span></span>|<span data-ttu-id="1cbaa-109">**类型**</span><span class="sxs-lookup"><span data-stu-id="1cbaa-109">**Type**</span></span>|<span data-ttu-id="1cbaa-110">**必需**</span><span class="sxs-lookup"><span data-stu-id="1cbaa-110">**Required**</span></span>|<span data-ttu-id="1cbaa-111">**说明**</span><span class="sxs-lookup"><span data-stu-id="1cbaa-111">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="1cbaa-112">名称</span><span class="sxs-lookup"><span data-stu-id="1cbaa-112">Name</span></span>|<span data-ttu-id="1cbaa-113">字符串</span><span class="sxs-lookup"><span data-stu-id="1cbaa-113">string</span></span>|<span data-ttu-id="1cbaa-114">必需</span><span class="sxs-lookup"><span data-stu-id="1cbaa-114">required</span></span>|<span data-ttu-id="1cbaa-p101">指定由其父对象限定的所需方法的名称。例如，要指定 **getSelectedDataAsync** 方法，必须指定 `"Document.getSelectedDataAsync"`。</span><span class="sxs-lookup"><span data-stu-id="1cbaa-p101">Specifies the name of the required method qualified with its parent object. For example, to specify the  **getSelectedDataAsync** method, you must specify `"Document.getSelectedDataAsync"`.</span></span>|

## <a name="remarks"></a><span data-ttu-id="1cbaa-117">说明</span><span class="sxs-lookup"><span data-stu-id="1cbaa-117">Remarks</span></span>

<span data-ttu-id="1cbaa-118">邮件外接程序不支持的**方法**和**方法**的元素。有关要求集的详细信息，请参阅[Office 版本和要求集](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)。</span><span class="sxs-lookup"><span data-stu-id="1cbaa-118">The  **Methods** and **Method** elements aren't supported by mail add-ins. For more information about requirement sets, see [Office versions and requirement sets](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="1cbaa-119">由于无法指定单个方法的最低版本要求，以确保方法是在运行时，您应该还使用**if**语句的加载项在脚本中调用该方法时。</span><span class="sxs-lookup"><span data-stu-id="1cbaa-119">Because there is no way to specify the minimum version requirement for individual methods, to make sure that a method is available at runtime, you should also use an **if** statement when calling that method in the script of your add-in.</span></span> <span data-ttu-id="1cbaa-120">有关如何执行此操作的详细信息，请参阅[Understanding the JavaScript API for Office](https://docs.microsoft.com/office/dev/add-ins/develop/understanding-the-javascript-api-for-office)。</span><span class="sxs-lookup"><span data-stu-id="1cbaa-120">For more information about how to do this, see [Understanding the JavaScript API for Office](https://docs.microsoft.com/office/dev/add-ins/develop/understanding-the-javascript-api-for-office).</span></span>

