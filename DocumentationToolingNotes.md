# <a name="how-the-office-javascript-api-documentation-is-generated"></a>如何生成 Office JavaScript API 文档

Office JavaScript 参考文档页从类型定义文件和示例代码段生成。 此过程混合使用开放源代码工具和特定于存储库的脚本。 本文档旨在使此存储库的流程透明，以便社区可以更好地受益于此内容并为此内容做贡献。

## <a name="content-sources"></a>内容源

组合了两种类型的内容来创建 Office-JS 参考文档：类型定义和代码段。 这些示例可确保完整的 API 覆盖范围，并提供了小型的内嵌代码示例。

### <a name="type-definition-files"></a>类型定义文件

"一 [定类型"](https://github.com/DefinitelyTyped/DefinitelyTyped) 上的类型定义文件是文档唯一真实来源。 使用 TypeScript 的任何 Office 加载项都使用这些类型定义文件进行编译。 这还会为 JavaScript 和 TypeScript 开发人员IntelliSense功能。 通过根据这些定义生成参考文档，我们提供了更准确的信息。

有四个相关的 d.ts 文件，它们为文档的不同子节提供源内容。

- [office-js/index.d.ts](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js/index.d.ts) (Release definitions.) 
  - [Excel (发布) ](https://docs.microsoft.com/javascript/api/excel_release)
  - [OneNote](https://docs.microsoft.com/javascript/api/onenote)
  - [PowerPoint](https://docs.microsoft.com/javascript/api/powerpoint)
  - [Visio](https://docs.microsoft.com/javascript/api/visio)
  - [Word (发布) ](https://docs.microsoft.com/javascript/api/word_release)
  - [通用 API 的 OfficeExtensions 子节](https://docs.microsoft.com/javascript/api/office)
- [office-js-preview/index.d.ts](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts) (Preview definitions.) 
  - [Excel (预览) ](https://docs.microsoft.com/javascript/api/excel)
  - [Outlook (预览) ](https://docs.microsoft.com/javascript/api/outlook)
  - [Word (预览) ](https://docs.microsoft.com/javascript/api/word)
  - [通用 API](https://docs.microsoft.com/javascript/api/office)
- [custom-functions-runtime/index.d.ts](https://github.com/DefinitelyTyped/DefinitelyTyped/blob/master/types/custom-functions-runtime/index.d.ts) (Excel Custom Functions 运行时定义.) 
  - [自定义函数](https://docs.microsoft.com/javascript/api/custom-functions-runtime)
- [office-runtime/index.d.ts](https://github.com/DefinitelyTyped/DefinitelyTyped/blob/master/types/office-runtime/index.d.ts) (Custom Functions platform.) 
  - [Office 运行时](https://docs.microsoft.com/javascript/api/office-runtime)

较旧版本的 API 具有其自己的 d.ts 文件。 当发布新的 API 要求集时，将保留这些要求。 它们也可使用版本删除 [程序工具生成](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/tools/VersionRemover.ts)。 这些旧的 d.ts 文件会进行维护，以便对 API 进行修补或更改时，仍记录原始行为。 如果你必须面向较早版本的 API，这将非常有用。

#### <a name="testing-type-definition-file-changes"></a>测试类型定义文件更改

Office JavaScript API 的任何文档更改都通过编辑上述四个 d.ts 文件完成。 但是，如果需要，例如，在 [generate-docs/script-inputs](https://github.com/OfficeDev/office-js-docs-reference/tree/master/generate-docs/script-inputs) 中编辑相应的文件并运行 [GenerateDocs.cmd，](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/GenerateDocs.cmd)可以在将 PR 提交到 DefinitelyTyped (之前测试更改。例如，测试格式如何转换为 markdown) 。 当系统提示时，选择"本地文件"选项。

将更改推送到此存储库的远程分支将导致docs.microsoft.com构建测试分支。 此分支在 review.docs.microsoft.com 上呈现，只有内部 Microsoft 人员可以访问它。 任何查看 PR 的人都将检查评价网站的准确性。

### <a name="code-snippets"></a>代码段

代码示例代码段从两个源添加到引用页面：

- [Script Lab 示例](https://github.com/OfficeDev/office-js-snippets)
- [本地代码段](https://github.com/OfficeDev/office-js-docs-reference/tree/master/docs/code-snippets)

本地代码段在主机特定的 yaml 文件中。 其内容按类和字段进行组织，因此可以映射到引用页中的适当位置。 JavaScript 或 TypeScript (代码段) 通过使用 await 语句推断。

脚本实验室代码段从工作示例提取。 目前，Excel、Outlook、PowerPoint 和 Word 示例通过映射文件映射到引用 [文档节](https://github.com/OfficeDev/office-js-snippets/tree/prod/snippet-extractor-metadata)。 这些示例方法将单个示例方法与 API 中的属性或方法匹配。 当 office-js-snippets 存储库运行时，将创建一个包含所有映射代码段 `yarn start` 的 [yaml](https://github.com/OfficeDev/office-js-snippets/blob/prod/snippet-extractor-output/snippets.yaml) 文件。 此 yaml 文件是参考文档工具的输入。

## <a name="tooling-pipeline"></a>工具管道

![显示从"明确类型"到预处理器、API 提取程序、中间处理器、API 文档器以及后处理器的控件流的图像。](ToolingPipeline.png)

在内容源和最终页面之间，文档内容将执行五个工具步骤：

1. [预处理器脚本](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/scripts/preprocessor.ts)
1. [API 提取程序](https://api-extractor.com/)
1. [Midprocessor 脚本](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/scripts/midprocessor.ts)
1. [API 文档器](https://github.com/microsoft/rushstack/blob/master/apps/api-documenter/README.md)
1. [后处理器脚本](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/scripts/postprocessor.ts)

预处理器会采用 d.ts 文件，并拆分为主机特定的部分。 它将执行后续工具正确处理数据所需的任何清理。

API 提取程序将 d.ts 文件转换为 JSON 数据。 这将标记所有类型数据，从而简化分析。

中间处理器检索代码段，将它们与正确的主机配对，并清理 Outlook 和通用 API 对象之间的交叉链接。

API 文档程序将 JSON 数据转换为 .yml 文件。 .yml 文件由将文档发布到文档发布系统的 Open Publishing System 转换为 markdown docs.microsoft.com。 API 文档器还包含插入代码段的特定于 Office 的扩展。

后置处理程序清理目录，将 .yml 文件移动到 [发布文件夹](https://github.com/OfficeDev/office-js-docs-reference/tree/master/docs/docs-ref-autogen)。

所有五个步骤在 [GenerateDocs.cmd](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/GenerateDocs.cmd) 运行时执行。 该脚本还处理节点模块安装、清除旧文件集以及每个要求集的版本类型定义文件。
