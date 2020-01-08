# <a name="how-the-office-javascript-api-documentation-is-generated"></a>Office JavaScript API 文档的生成方式

Office JavaScript 参考文档页面由类型定义文件和示例代码段生成。 此过程使用开放源工具和存储库特定的脚本的混合。 本文档旨在使此存储库的过程透明，以便社区能够更好地利用和参与此内容。

## <a name="content-sources"></a>内容源

组合两种类型的内容以创建 Office JS 参考文档：类型定义和代码段。 这些确保了完整的 API 覆盖率并提供了小型的内联代码示例。

### <a name="type-definition-files"></a>类型定义文件

[绝对键入](https://github.com/DefinitelyTyped/DefinitelyTyped)的类型定义文件是文档的唯一事实来源。 使用这些类型定义文件的使用 TypeScript 编译的任何 Office 外接程序。 这些功能还提供 JavaScript 和 TypeScript 开发人员 IntelliSense 功能。 通过生成这些定义中的参考文档，我们可以提供更准确的信息。

共有四个相关的 d ts 文件，它们提供文档不同子部分的源内容。

- [office-js/索引](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js/index.d.ts)（发布定义）。
  - [Excel （发布）](https://docs.microsoft.com/javascript/api/excel_release)
  - [OneNote](https://docs.microsoft.com/javascript/api/onenote)
  - [PowerPoint](https://docs.microsoft.com/javascript/api/powerpoint)
  - [Visio](https://docs.microsoft.com/javascript/api/visio)
  - [Word （发布）](https://docs.microsoft.com/javascript/api/word_release)
  - [通用 API 的 OfficeExtensions 子部分](https://docs.microsoft.com/javascript/api/office)
- [office-js-预览/索引](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts)（预览定义）。
  - [Excel （预览）](https://docs.microsoft.com/javascript/api/excel)
  - [Outlook （预览）](https://docs.microsoft.com/javascript/api/outlook)
  - [Word （预览）](https://docs.microsoft.com/javascript/api/word)
  - [通用 API](https://docs.microsoft.com/javascript/api/office)
- [自定义函数-运行时](https://github.com/DefinitelyTyped/DefinitelyTyped/blob/master/types/custom-functions-runtime/index.d.ts)（Excel 自定义函数运行时定义。）
  - [自定义函数](https://docs.microsoft.com/javascript/api/custom-functions-runtime)
- [office 运行时](https://github.com/DefinitelyTyped/DefinitelyTyped/blob/master/types/office-runtime/index.d.ts)（自定义函数平台的 office 运行时定义。）
  - [Office 运行时](https://docs.microsoft.com/javascript/api/office-runtime)

较旧版本的 Api 具有它们自己的 d ts 文件。 当发布新的 API 要求集时，将保留这些设置。 也可以使用[版本 Remover 工具](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/tools/VersionRemover.ts)生成。 保留这些旧的 d .aspx 文件，以便在对事件 Api 进行修补或更改时，仍会记录原始行为。 如果您必须针对较旧版本的 API，这将非常有用。

### <a name="code-snippets"></a>代码段

代码示例代码段将添加到以下两个源中的参考页面：

- [脚本实验室示例](https://github.com/OfficeDev/office-js-snippets)
- [本地代码段](https://github.com/OfficeDev/office-js-docs-reference/tree/master/docs/code-snippets)

本地代码段位于特定于主机的 yaml 文件中。 其内容按类和字段进行组织，以便可以将其映射到参考页中的适当位置。 使用 await 语句可推断代码段（JavaScript 或 TypeScript）的语言。

脚本实验室代码段是从工作示例中提取的。 目前，Excel 和 Word 示例通过一[对映射文件](https://github.com/OfficeDev/office-js-snippets/tree/master/snippet-extractor-metadata)映射到参考文档部分。 这些方法将各个示例方法与 API 中的属性或方法相匹配。 在`yarn start`运行 office js 存储库时，将创建一个包含所有已映射代码段的[yaml 文件](https://github.com/OfficeDev/office-js-snippets/blob/master/snippet-extractor-output/snippets.yaml)。 此 yaml 文件是参考文档工具中的输入。

## <a name="tooling-pipeline"></a>刀具管道

![显示从绝对类型到预处理器、API 提取程序、midprocessor、API 文档管理和到 postprocessor 的控制流的图像。](ToolingPipeline.png)

在内容源和最终页面之间，文档内容将经历五个工具步骤：

1. [预处理器脚本](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/scripts/preprocessor.ts)
1. [API 提取程序](https://api-extractor.com/)
1. [Midprocessor 脚本](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/scripts/midprocessor.ts)
1. [API 文档](https://github.com/microsoft/rushstack/blob/master/apps/api-documenter/README.md)
1. [Postprocessor 脚本](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/scripts/postprocessor.ts)

预处理器采用 d. ts 文件，并将它们拆分为特定于主机的分区。 它执行后续工具正确处理数据所需的任何清理操作。

API 提取程序将 d. ts 文件转换为 JSON 数据。 此 tokenizes 所有类型数据，从而便于进行分析。

Midprocessor 检索代码段并将它们与适当的主机配对。

API 文档管理器将 JSON 数据转换为 yml 文件。 将文档发布到 docs.microsoft.com 的开放发布系统会将. yml 文件转换为 markdown。 API 文档管理器还包含一个特定于 Office 的扩展，用于插入代码段。

Postprocessor 清理目录并将. yml 文件移到[发布文件夹](https://github.com/OfficeDev/office-js-docs-reference/tree/master/docs/docs-ref-autogen)中。

在运行[GenerateDocs](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/GenerateDocs.cmd)时，将执行上述所有五个步骤。 该脚本还处理节点模块安装、清除旧文件集和每个要求集的版本类型定义文件。
