# <a name="javascript-api-for-office"></a>适用于 Office 的 JavaScript API

JavaScript API for Office 使您能够与 Office 主机应用程序中的对象模型创建的 web 应用程序进行交互。 您的应用程序将引用 office.js 库，它是脚本加载程序。 在 office.js 库加载适用于 Office 应用程序正在运行的加载项对象模型。 您可以使用以下 JavaScript 对象模型：

- **公共 Api** -引入使用**Office 2013**Api。 这是针对**所有 Office 主机应用程序**加载，与 Office 客户端应用程序连接外接程序应用程序。 对象模型中包含的特定于 Office 客户端 Api 和适用于多个 Office 客户端主机应用程序的 Api。 此内容的所有正在**共享的 API**。 

  **Outlook**也使用的常用语法都 API。 别名 Office 下的所有内容包含可用于编写脚本以便与 Office 文档、 工作表、 演示文稿、 邮件项目和从您 Office 加载项项目中的内容进行交互的对象。如果加载项将面向 Office 2013 和更高版本，则必须使用这些公共 Api。 此对象模型使用回调。

- **特定于宿主的 Api** -与**Office 2016**引入 Api。 此对象模型提供了特定于宿主的强类型对象对应于时使用 Office 客户端和表示 Office JavaScript Api 的发展，您看到的熟悉对象。 特定于宿主的 Api 当前包括 Word 的 JavaScript API 和 Excel 的 JavaScript API。

## <a name="supported-host-applications"></a>支持的主机应用程序

- [Excel](overview/excel-add-ins-reference-overview.md)
- [OneNote](overview/onenote-add-ins-javascript-reference.md)
- [Outlook](requirement-sets/outlook-api-requirement-sets.md)
- [Visio](overview/visio-javascript-reference-overview.md)
- [Word](overview/word-add-ins-reference-overview.md)
- [共享 API](requirement-sets/office-add-in-requirement-sets.md)

> [!NOTE] 
> [PowerPoint 和 Project](requirement-sets/powerpoint-and-project-note.md)支持使用 JavaScript API 所做的加载项。 但是，他们当前不具有特定主机的 Api。 您与这些主机通过共享 API 进行交互。

了解有关[支持的主机和其他要求](https://docs.microsoft.com/office/dev/add-ins/concepts/requirements-for-running-office-add-ins)的详细信息。

## <a name="open-api-specifications"></a>开放 API 规范

在我们设计和开发新的 API 以用于 Office 外接程序时，我们将使它们适用于[开放 API 规范](openspec.md)页的反馈。了解管道中的新增功能，并提供您对我们的设计规范的宝贵意见。

## <a name="see-also"></a>另请参阅

- [Office 的 JavaScript API 参考 （英文）](https://docs.microsoft.com/javascript/api/overview/office?view=office-js)