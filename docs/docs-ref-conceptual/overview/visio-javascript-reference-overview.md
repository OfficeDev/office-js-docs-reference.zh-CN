# <a name="visio-javascript-api-overview"></a>Visio 的 JavaScript API 概述

可以使用 Visio JavaScript API 在 SharePoint Online 中嵌入 Visio 图表。 嵌入的 Visio 图表是存储在 SharePoint 文档库并在 SharePoint 页面上显示的图表。 若要嵌入的 Visio 图表，显示在 HTML`<iframe>`元素。 然后，可以使用 Visio JavaScript API 以程序化方式处理嵌入的图表。

![SharePoint 页面上 iframe 中的 Visio 图表，以及脚本编辑器 Web 部件](../images/visio-api-block-diagram.png)


可以使用 Visio JavaScript API 执行以下操作：

* 与 Visio 图表元素，如页面和形状进行交互。
* 在 Visio 图示画布上创建可视标记。
* 编写绘图中的鼠标事件的自定义处理的程序。
* 向解决方案公开图表数据，如形状文本、形状数据和超链接。

本文介绍了如何通过结合使用 Visio JavaScript API 和 Visio Online 来生成 SharePoint Online 解决方案。具体介绍了有关使用 API（如 **EmbeddedSession**、**RequestContext**、JavaScript 代理对象、**sync()**、**Visio.run()** 和 **load()** 方法）的基本概念。下面这些代码示例展示了如何应用这些概念。

## <a name="embeddedsession"></a>EmbeddedSession

EmbeddedSession 对象初始化开发者框架和 Visio Online 框架之间的通信。

```js
var session = new OfficeExtension.EmbeddedSession(url, { id: "embed-iframe",container: document.getElementById("iframeHost") });
session.init().then(function () {    
    window.console.log("Session successfully initialized");
});
```

## <a name="visiorunsession-functioncontext--batch-"></a>Visio.run （会话，function(context) {批处理}）

**Visio.run()** 运行一个对 Visio 对象模型执行操作的批处理脚本。 批处理命令包括定义本地 JavaScript 代理对象、在本地和 Visio 对象之间同步状态的 **sync()** 方法以及承诺实现。 **Visio.run()** 中的批处理请求的优势在于，当实现承诺时，在执行期间分配的任何被跟踪的页面对象将会自动释放。 

运行的方法接受会话和 RequestContext 对象并返回承诺 (通常刚刚**context.sync()** 的结果)。 可以在 **Visio.run()** 之外运行批处理操作。 不过，在这种情况下，需要手动跟踪和管理任何页面对象引用。 

## <a name="requestcontext"></a>RequestContext

RequestContext 对象方便了 Visio 应用程序的请求。 开发人员图文框与 Visio Online 应用程序在两个不同的 iframe 中运行，因为需要 RequestContext 对象 （在下面的示例上下文） 来获得对 Visio 和相关的对象，如网页和形状，从开发人员框架的访问。 

```js
function hideToolbars() {
    Visio.run(session, function(context){
        var app = context.document.application;
        app.showToolbars = false;            
        return context.sync().then(function ()
        {
            window.console.log("Toolbars Hidden");
        });      
        }).catch(function(error)
    {
        window.console.log("Error: " + error);            
    });
};
```

## <a name="proxy-objects"></a>代理对象

在外接程序中声明和使用的 Visio JavaScript 对象是 Visio 文档中真实对象的代理对象。对代理对象执行的所有操作都不会在 Visio 中实现；在同步文档状态前，Visio 文档的状态不会在代理对象中实现。运行 `context.sync()` 时将同步文档状态。

例如，本地 JavaScript 对象 getActivePage 声明引用所选的页。 这可以用于对其属性和调用方法的设置进行排队。 在运行**sync()** 方法之前不实现此类对象上的操作。

```js
var activePage = context.document.getActivePage();
```

## <a name="sync"></a>Sync

**Sync()** 方法同步 JavaScript 代理对象之间的状态和 Visio 中的实际对象，通过执行说明排队的上下文的检索属性加载在您的代码中使用的 Office 对象。 此方法返回一个将在同步完成时实现的承诺。 

## <a name="load"></a>load()

**load()** 方法用于填充在外接程序 JavaScript 层中创建的代理对象。尝试检索对象（如文档）时，将首先在 JavaScript 层中创建一个本地代理对象。此类对象可用于将属性和调用方法的设置操作排入队列。不过，若要读取对象属性或关系，必须先调用 **load()** 和 **sync()** 方法。load() 方法获取在调用 **sync()** 方法时需要加载的属性和关系。

下面的示例展示了 **load()** 方法的语法。

```js
object.load(string: properties); //or object.load(array: properties); //or object.load({loadOption});
```

1. **属性**是属性名称为加载指定的以逗号分隔字符串的列表或名称的数组。 有关详细信息，请参阅每个对象下的 **.load()** 方法。

2. **loadOption** 指定的对象描述了选择、展开、置顶和跳过选项。有关详细信息，请参阅对象加载[选项](/javascript/api/office/officeextension.loadoption)。

## <a name="example-printing-all-shapes-text-in-active-page"></a>示例：打印活动页中的所有形状文本

下面的示例展示了如何打印数组形状对象的形状文本值。 **Visio.run()** 方法包含一批指令。 在此次批处理期间，将会创建一个代理对象，引用活动文档中的形状。

所有这些命令是排队，并在调用**context.sync()** 时运行。 **sync()** 方法返回一个承诺，可用于将其与其他操作关联起来。

```js
Visio.run(session, function (context) {
   var page = context.document.getActivePage();
   var shapes = page.shapes;
   shapes.load();
   return context.sync().then(function () {
        for(var i=0; i<shapes.items.length;i++)
 {
            var shape = shapes.items[i];
     window.console.log("Shape Text: " + shape.text );
 }
});
}).catch(function(error) {
  window.console.log("Error: " + error);
  if (error instanceof OfficeExtension.Error) {
       window.console.log ("Debug info: " + JSON.stringify(error.debugInfo));
  }
});
```

## <a name="error-messages"></a>错误消息

使用包含代码和消息的错误对象返回错误。下表列出了可能发生的错误情况。

| error.code            | error.message |
|-----------------------|----------------------------------------------------------------|
| InvalidArgument       | 自变量无效、缺少或格式不正确。 |
| GeneralException      | 处理请求时出现内部错误。 |
| NotImplemented        | 所请求的功能未实现。  |
| UnsupportedOperation  | 不支持正在尝试的操作。 |
| AccessDenied          | 无法执行所请求的操作。 |
| ItemNotFound          | 所请求的资源不存在。 |

## <a name="get-started"></a>开始使用

您可以使用本节中的示例开始。 本示例演示如何以编程方式将 Visio 图表中显示选定形状的形状文本。 若要开始，在 SharePoint Online 中创建经典页或编辑现有网页。 页面上添加脚本编辑器 web 部件和复制和粘贴以下代码。

```js
<script src='https://appsforoffice.microsoft.com/embedded/1.0/visio-web-embedded.js' type='text/javascript'></script>

Enter Visio File Url:<br/>
<script language="javascript">
document.write("<input type='text' id='fileUrl' size='120'/>");
document.write("<input type='button' value='InitEmbeddedFrame' onclick='initEmbeddedFrame()' />");
document.write("<br />");
document.write("<input type='button' value='SelectedShapeText' onclick='getSelectedShapeText()' />");
document.write("<textarea id='ResultOutput' style='width:350px;height:60px'> </textarea>");
document.write("<div id='iframeHost' />");

let session; // Global variable to store the session and pass it afterwards in Visio.run()
var textArea;
// Loads the Visio application and Initializes communication between devloper frame and Visio online frame
function initEmbeddedFrame() {
        textArea = document.getElementById('ResultOutput');
    var url = document.getElementById('fileUrl').value;
    if (!url) {
        window.alert("File URL should not be empty");
    }
    // APIs are enabled for EmbedView action only.   
    url = url.replace("action=view","action=embedview");
    url = url.replace("action=interactivepreview","action=embedview");
    url = url.replace("action=default","action=embedview");
    url = url.replace("action=edit","action=embedview");
  
       session = new OfficeExtension.EmbeddedSession(url, { id: "embed-iframe",container: document.getElementById("iframeHost") });
       return session.init().then(function () {
        // Initilization is successful 
        textArea.value  = "Initilization is successful";
    });
     }

// Code for getting selected Shape Text using the shapes collection object
function getSelectedShapeText() {
    Visio.run(session, function (context) {     
       var page = context.document.getActivePage();
       var shapes = page.shapes;
       shapes.load();
           return context.sync().then(function () {
               textArea.value = "Please select a Shape in the Diagram";
               for(var i=0; i<shapes.items.length;i++)
            {
              var shape = shapes.items[i];
                  if ( shape.select == true)
               {
                textArea.value = shape.text;
                    return;
                   }
            }
      });
     }).catch(function(error) {
        textArea.value = "Error: ";
        if (error instanceof OfficeExtension.Error) {
            textArea.value += "Debug info: " + JSON.stringify(error.debugInfo);
        }
    });
}
</script>
```

之后，所需的只是您想要使用的 Visio 图表的 URL。 只需将 Visio 图表上载到 SharePoint Online，然后打开该 Visio Online 中。 从该处，打开嵌入对话框，并在上面的示例使用嵌入的 URL。

![将 Visio 文件 URL 复制从嵌入对话框](../images/Visio-embed-url.png)

如果您使用 Visio Online 在编辑模式下，打开嵌入对话框中，请选择**文件** > **共享** > **嵌入**。 如果您使用 Visio 联机视图模式中，打开嵌入对话框中，请选择...，然后**嵌入**。 

## <a name="open-api-specifications"></a>开放 API 规范

在设计和开发新的 API 时，我们会“[开放性 API 规范](../openspec.md)”页面上提供这些 API，以便你向我们提供反馈。了解管道中的新增功能，并提供你对我们的设计规范的宝贵意见。 

## <a name="visio-javascript-api-reference"></a>Visio 的 JavaScript API 参考 （英文）

有关 Visio JavaScript API 的详细信息，请参阅[Visio 的 JavaScript API 参考文档](/javascript/api/visio)。
