# <a name="excel-javascript-api-requirement-sets"></a>Excel JavaScript API 要求集

要求集是指各组已命名的 API 成员。 Office 加载项使用清单中指定要求集，或使用运行时检查以确定是否的 Office 主机支持外接程序需要的 Api。 有关详细信息，请参阅[Office 版本和要求集](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)。

Excel 加载项运行跨多个版本的 Office，包括 Office 2016 或更高版本的 Windows、 ipad 版的 Office、 Office for Mac，和 Office Online。 下表列出了 Excel 要求集，支持每个要求集和生成版本或数字这些应用程序的 Office 主机应用程序。

> [!NOTE]
> 标记为**Beta**任何 API 不为最终用户生产做好准备。 我们使其可供开发人员执行试用测试和开发环境中。 它们不旨在用于针对生产/业务关键文档。
> 
> 对于要求集标记为**Beta**，使用指定 （或更高版本） 版的 Office 软件和使用 CDN Beta 库： https://appsforoffice.microsoft.com/lib/beta/hosted/office.js。 未标记为**Beta**的条目是否通常可用，以及您可以使用生产库上 CDN: https://appsforoffice.microsoft.com/lib/1/hosted/office.js。

|  要求集  |  用于 Windows 的 office 365\*  |  Office 365 for iPad  |  Office 365 for Mac  | Office Online  |  Office Online Server  |
|:-----|-----|:-----|:-----|:-----|:-----|
| Beta  | 请[访问我们的 Excel 的 JavaScript API 打开规范页面](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec)！ |
| ExcelApi1.8  | 版本 1808 (生成 10730.20102) 或更高版本 | 2.17 或更高版本 | 16.17 或更高版本 | 年 9 月 2018 | 即将推出 |
| ExcelApi1.7  | 版本 1801 (生成 9001.2171) 或更高版本   | 2.9 或更高版本 | 16.9 或更高版本 | 2018 年 4 月 | 即将推出 |
| ExcelApi1.6  | 版本 1704（生成号 8201.2001）或更高版本   | 2.2 或更高版本 |15.36 或更高版本| 2017 年 4 月 | 即将推出|
| ExcelApi1.5  | 版本 1703（内部版本 8067.2070）或更高版本   | 2.2 或更高版本 |15.36 或更高版本| 2017 年 3 月 | 即将推出|
| ExcelApi1.4  | 版本 1701（内部版本 7870.2024）或更高版本   | 2.2 或更高版本 |15.36 或更高版本| 2017 年 1 月 | 即将推出|
| ExcelApi1.3  | 版本 1608（内部版本 7369.2055）或更高版本 | 1.27 或更高版本 |  15.27 或更高版本| 2016 年 9 月 | 版本 1608（内部版本 7601.6800）或更高版本|
| ExcelApi1.2  | 版本 1601（内部版本 6741.2088）或更高版本 | 1.21 或更高版本 | 15.22 或更高版本| 2016 年 1 月 ||
| ExcelApi1.1  | 版本 1509（内部版本 4266.1001）或更高版本 | 1.19 或更高版本 | 15.20 或更高版本| 2016 年 1 月 ||

> [!NOTE]
> 您可以通过 MSI 安装的 Office 2016 内部版本号是 16.0.4266.1001。 此版本中仅包含 ExcelApi 1.1 要求集。

有关版本、 内部版本号和 Office Online 服务器的详细信息，请参阅：

- [更新频道发布的 Office 365 客户端版本号和内部版本号](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [使用的是哪一版 Office？](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [在哪里可以找到 Office 365 客户端应用程序的版本号和内部版本号](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- 
  [Office Online Server 概述](https://docs.microsoft.com/officeonlineserver/office-online-server-overview)

## <a name="whats-new-in-excel-javascript-api-18"></a>What's new in Excel JavaScript API 1.8

Excel 的 JavaScript API 要求集 1.8 功能包括 Api 的数据透视表、 数据验证、 图表、 图表、 性能选项和工作簿创建的事件。

### <a name="pivottable"></a>PivotTable

波形 2 的数据透视表 Api 允许加载项设置的数据透视表的层次结构。 您现在可以控制数据和聚合方式。 我们的[数据透视表文章](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-pivottables)包含有关新的数据透视表功能的详细信息。

### <a name="data-validation"></a>数据验证

您可以控制哪些用户的进入工作表中的数据验证提供。 您可以限制预定义的应答设置单元格或提供有关不需要输入弹出警告。 了解有关[添加到区域的数据验证](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-data-validation)今天。

### <a name="charts"></a>图表

另一轮图表 Api 将图表元素的更大编程控制。 您现在可以更高版本访问图例、 轴、 趋势线和绘图区。

### <a name="events"></a>事件

对于图表，已经添加多个[事件](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-events)。 具有与图表交互的用户您外接程序做出反应。 您还可以[切换事件](https://docs.microsoft.com/office/dev/add-ins/excel/performance#enable-and-disable-events)触发跨整个工作簿。


|对象| 最近更新| 说明|要求集|
|:----|:----|:----|:----|
|[application](/javascript/api/excel/excel.application)|_方法_ > [createWorkbook(base64File: string)](/javascript/api/excel/excel.application)|使用可选的 base64 编码的.xlsx 文件创建一个新的隐藏工作簿。|1.8|
|[basicDataValidation](/javascript/api/excel/excel.basicdatavalidation)|_属性_> formula1|获取或设置 Formula1，即最小值或运算符的具体的值。|1.8|
|[basicDataValidation](/javascript/api/excel/excel.basicdatavalidation)|_属性_> formula2|获取或设置 Formula2，即最大值或运算符的具体的值。|1.8|
|[basicDataValidation](/javascript/api/excel/excel.basicdatavalidation)|_关系_> 运算符|要用于验证的数据的运算符。|1.8|
|[chart](/javascript/api/excel/excel.chart)|_属性_> categoryLabelLevel|返回或设置引用的其中源自分类标签正在级别 ChartCategoryLabelLevel 枚举常量。 读/写。|1.8|
|[chart](/javascript/api/excel/excel.chart)|_属性_> plotVisibleOnly|如此 如果只绘制可见单元格。 假 如果可见和隐藏单元格的绘制。 ReadWrite。|1.8|
|[chart](/javascript/api/excel/excel.chart)|_属性_> seriesNameLevel|返回或设置引用的其中源自系列名称正在级别 ChartSeriesNameLevel 枚举常量。 读/写。|1.8|
|[chart](/javascript/api/excel/excel.chart)|_属性_> showDataLabelsOverMaximum|代表时是否显示数据标签的值大于上数值轴的最大值。|1.8|
|[chart](/javascript/api/excel/excel.chart)|_属性_ > style|返回或设置图表的图表样式。 ReadWrite。|1.8|
|[chart](/javascript/api/excel/excel.chart)|_关系_> displayBlanksAs|返回或设置在图表上绘制空白单元格方式。 ReadWrite。|1.8|
|[chart](/javascript/api/excel/excel.chart)|_关系_> plotArea|代表图表 plotArea。 只读。|1.8|
|[chart](/javascript/api/excel/excel.chart)|_关系_> plotBy|返回或设置行或列用作数据系列在图表上的方式。 ReadWrite。|1.8|
|[chartActivatedEventArgs](/javascript/api/excel/excel.chartactivatedeventargs)|_属性_> chartId|获取的图表的激活 id。|1.8|
|[chartActivatedEventArgs](/javascript/api/excel/excel.chartactivatedeventargs)|_属性_> 类型|获取该事件的类型。|1.8|
|[chartActivatedEventArgs](/javascript/api/excel/excel.chartactivatedeventargs)|_属性_> worksheetId|获取在其中激活图表工作表的 id。|1.8|
|[chartAddedEventArgs](/javascript/api/excel/excel.chartaddedeventargs)|_属性_> chartId|获取添加到表中的图表的 id。|1.8|
|[chartAddedEventArgs](/javascript/api/excel/excel.chartaddedeventargs)|_属性_> 类型|获取该事件的类型。|1.8|
|[chartAddedEventArgs](/javascript/api/excel/excel.chartaddedeventargs)|_属性_> worksheetId|获取在其中添加图表工作表的 id。|1.8|
|[chartAddedEventArgs](/javascript/api/excel/excel.chartaddedeventargs)|_关系_> 源|获取在事件的源。|1.8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_属性_> isBetweenCategories|表示是否数值轴与分类坐标轴相交于分类之间。|1.8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_属性_> 多级|代表坐标轴是否多级。|1.8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_属性_ > numberFormat|代表坐标轴刻度线标签的格式代码。|1.8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_属性_> 偏移|代表标签的级别和第一级与坐标轴之间的距离之间的距离。 值应为从 0 到 1000 之间的整数。|1.8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_属性_> positionAt|代表与其他坐标轴相交在指定的坐标轴位置。 应使用 SetPositionAt(double) 方法设置该属性。 只读。|1.8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_属性_> textOrientation|代表坐标轴刻度线标签的文本方向。 值应为整数从-90 至 90 或 180 纵向文本。|1.8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_关系_> 对齐方式|代表指定的坐标轴刻度线标签的对齐方式。|1.8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_关系_> 位置|代表与其他坐标轴相交的指定的坐标轴位置。|1.8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_方法_ > [setPositionAt(value: double)](/javascript/api/excel/excel.chartaxis)|设置与其他坐标轴相交在指定的坐标轴位置。|1.8|
|[chartAxisFormat](/javascript/api/excel/excel.chartaxisformat)|_关系_> 填充|代表图表填充格式。 只读。|1.8|
|[chartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|_方法_ > [setFormula(formula: string)](/javascript/api/excel/excel.chartaxistitle)|一个字符串值，它代表图表坐标轴标题使用 A1 样式表示法的公式。|1.8|
|[chartAxisTitleFormat](/javascript/api/excel/excel.chartaxistitleformat)|_关系_> 边框|代表边框格式，其中包括颜色、 linestyle 和权重。 只读。|1.8|
|[chartAxisTitleFormat](/javascript/api/excel/excel.chartaxistitleformat)|_关系_> 填充|代表图表填充格式。 只读。|1.8|
|[chartBorder](/javascript/api/excel/excel.chartborder)|_方法_ > [clear()](/javascript/api/excel/excel.chartborder)|清除图表元素的边框格式。|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_属性_> 自动图文集|布尔值，表示如果数据标签自动生成相应基于上下文的文字。|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_属性_> 公式|String 值，它代表 A1 样式表示法的图表数据标签的公式。|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_属性_ > height|返回高度，以磅为单位的图表数据标签。 只读。 如果图表数据标签不可见，则为 null。 只读。|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_属性_> 左侧|从左边缘到图表区的左边缘的图表数据标签表示的距离，以磅为单位。 如果图表数据标签不可见，则为 null。|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_属性_ > numberFormat|String 值，它代表数据标签的格式代码。|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_属性_ > text|表示图表上的数据标签的文本字符串。|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_属性_> textOrientation|代表图表数据标签的文本方向。 值应为整数从-90 至 90 或 180 纵向文本。|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_属性_> 顶部|代表从图表的图表区域顶部的数据标签的上边缘的距离，以磅为单位。 如果图表数据标签不可见，则为 null。|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_属性_ > width|返回宽度，以磅为单位的图表数据标签。 只读。 如果图表数据标签不可见，则为 null。 只读。|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_关系_> 格式|代表图表数据标签的格式。 只读。|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_关系_ > horizontalAlignment|代表图表数据标签的水平对齐方式。|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_关系_ > verticalAlignment|代表图表数据标签的垂直对齐方式。|1.8|
|[chartDataLabelFormat](/javascript/api/excel/excel.chartdatalabelformat)|_关系_> 边框|代表边框格式，其中包括颜色、 linestyle 和权重。 只读。|1.8|
|[chartDataLabels](/javascript/api/excel/excel.chartdatalabels)|_属性_> 自动图文集|表示是否数据标签自动生成合适基于上下文的文字。|1.8|
|[chartDataLabels](/javascript/api/excel/excel.chartdatalabels)|_属性_ > numberFormat|代表数据标签的格式代码。|1.8|
|[chartDataLabels](/javascript/api/excel/excel.chartdatalabels)|_属性_> textOrientation|代表数据标签的文本方向。 值应为整数，一种从-90 至 90，或 0 到 180 纵向文本。|1.8|
|[chartDataLabels](/javascript/api/excel/excel.chartdatalabels)|_关系_ > horizontalAlignment|代表图表数据标签的水平对齐方式。|1.8|
|[chartDataLabels](/javascript/api/excel/excel.chartdatalabels)|_关系_ > verticalAlignment|代表图表数据标签的垂直对齐方式。|1.8|
|[chartDeactivatedEventArgs](/javascript/api/excel/excel.chartdeactivatedeventargs)|_属性_> chartId|获取停用的图表的 id。|1.8|
|[chartDeactivatedEventArgs](/javascript/api/excel/excel.chartdeactivatedeventargs)|_属性_> 类型|获取该事件的类型。|1.8|
|[chartDeactivatedEventArgs](/javascript/api/excel/excel.chartdeactivatedeventargs)|_属性_> worksheetId|获取一个图表的被停用的工作表的 id。|1.8|
|[chartDeletedEventArgs](/javascript/api/excel/excel.chartdeletedeventargs)|_属性_> chartId|获取的图表工作表中删除的 id。|1.8|
|[chartDeletedEventArgs](/javascript/api/excel/excel.chartdeletedeventargs)|_属性_> 类型|获取该事件的类型。|1.8|
|[chartDeletedEventArgs](/javascript/api/excel/excel.chartdeletedeventargs)|_属性_> worksheetId|获取在其中删除图表工作表的 id。|1.8|
|[chartDeletedEventArgs](/javascript/api/excel/excel.chartdeletedeventargs)|_关系_> 源|获取在事件的源。|1.8|
|[chartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|_属性_ > height|代表图表图例 legendEntry 的高度。 只读。|1.8|
|[chartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|_属性_ > index|代表图表图例中 legendEntry 的索引。 只读。|1.8|
|[chartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|_属性_> 左侧|代表图表 legendEntry 的左侧。 只读。|1.8|
|[chartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|_属性_> 顶部|代表图表 legendEntry 顶部。 只读。|1.8|
|[chartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|_属性_ > width|代表图表图例上 legendEntry 的宽度。 只读。|1.8|
|[chartLegendFormat](/javascript/api/excel/excel.chartlegendformat)|_关系_> 边框|代表边框格式，其中包括颜色、 linestyle 和权重。 只读。|1.8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_属性_ > height|代表 plotArea 的高度值。|1.8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_属性_> insideHeight|代表 plotArea 的 insideHeight 值。|1.8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_属性_> insideLeft|代表 plotArea insideLeft 值。|1.8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_属性_> insideTop|代表 plotArea insideTop 值。|1.8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_属性_> insideWidth|代表 plotArea insideWidth 值。|1.8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_属性_> 左侧|代表 plotArea 左侧的值。|1.8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_属性_> 顶部|代表 plotArea 的顶部的值。|1.8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_属性_ > width|代表 plotArea 的宽度值。|1.8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_关系_> 格式|代表图表 plotArea 的格式。 只读。|1.8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_关系_> 位置|代表 plotArea 的位置。|1.8|
|[chartPlotAreaFormat](/javascript/api/excel/excel.chartplotareaformat)|_关系_> 边框|代表图表 plotArea 的边框属性。 只读。|1.8|
|[chartPlotAreaFormat](/javascript/api/excel/excel.chartplotareaformat)|_关系_> 填充|表示对象的填充格式，包括背景格式信息。只读。|1.8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_属性_> explosion|返回或设置饼图或圆环图的扇区的分离程度值。 如果没有分离 （切片的提示是与饼图中心中），则返回 0 （零）。 ReadWrite。|1.8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_属性_> firstSliceAngle|返回或设置从顺时针 （垂直） 中的第一个复合饼图或圆环图扇面的角度。 仅适用于饼图、 三维饼图和圆环图。 可以为 0 到 360 之间的值。 ReadWrite|1.8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_属性_> invertIfNegative|如果 Microsoft Excel 反转项中的图案，当它对应于为负数，则为 true。 ReadWrite。|1.8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_属性_> 重叠|指定条形和柱形的排列方式。 可以为-100 到 100 之间的值。 仅适用于二维条形图和二维柱形图。 ReadWrite。|1.8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_属性_> secondPlotSize|返回或设置的第二部分的大小复合饼图或复合条饼图，主要饼图的大小的百分比。 可以为从 5 到 200 的值。 ReadWrite。|1.8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_属性_> varyByCategories|如果 Microsoft Excel 将不同的颜色或模式分配给每个数据标记，则为 true。 该属性所应用的图表只能包含一个系列。 ReadWrite。|1.8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_关系_> axisGroup|返回或设置指定的数据系列的组。 ReadWrite|1.8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_关系_> dataLabels|代表系列中的所有数据标签的集合。 只读。|1.8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_关系_> splitType|返回或设置复合饼图或复合条饼图的两个部分将被拆分的方式。 ReadWrite。|1.8|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_属性_> backwardPeriod|代表趋势线向后延伸的周期数。|1.8|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_属性_> forwardPeriod|代表趋势线向前延伸的周期数。|1.8|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_属性_> showEquation|如果在图表上显示趋势线公式，则为 true。|1.8|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_属性_> showRSquared|如果 R 平方趋势线在图表上显示，则为 true。|1.8|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_关系_> 标签|代表图表中趋势线的标签。 只读。|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_属性_> 自动图文集|布尔值，表示如果趋势线标签自动生成相应基于上下文的文字。|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_属性_> 公式|String 值，它代表 A1 样式表示法的图表趋势线标签的公式。|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_属性_ > height|返回以磅为单位的图表趋势线标签的高度。 只读。 如果图表趋势线标签不可见，则为 null。 只读。|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_属性_> 左侧|代表中图表区的左边缘到图表趋势线标签的左边缘的距离，以磅为单位。 如果图表趋势线标签不可见，则为 null。|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_属性_ > numberFormat|String 值，它代表趋势线标签的格式代码。|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_属性_ > text|表示图表上的趋势线标签的文本字符串。|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_属性_> textOrientation|代表图表趋势线标签的文本方向。 值应为整数从-90 至 90 或 180 纵向文本。|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_属性_> 顶部|代表从图表的图表区域顶部的趋势线标签的上边缘的距离，以磅为单位。 如果图表趋势线标签不可见，则为 null。|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_属性_ > width|返回宽度，以磅为单位的图表趋势线标签。 只读。 如果图表趋势线标签不可见，则为 null。 只读。|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_关系_> 格式|代表图表趋势线标签的格式。 只读。|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_关系_ > horizontalAlignment|代表图表趋势线标签的水平对齐方式。|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_关系_ > verticalAlignment|代表图表趋势线标签的垂直对齐方式。|1.8|
|[chartTrendlineLabelFormat](/javascript/api/excel/excel.charttrendlinelabelformat)|_关系_> 边框|代表边框格式，其中包括颜色、 linestyle 和权重。 只读。|1.8|
|[chartTrendlineLabelFormat](/javascript/api/excel/excel.charttrendlinelabelformat)|_关系_> 填充|代表当前图表趋势线标签的填充格式。 只读。|1.8|
|[chartTrendlineLabelFormat](/javascript/api/excel/excel.charttrendlinelabelformat)|_关系_ > font|代表图表趋势线标签的字体属性 （字体名称、 字号、 颜色等）。 只读。|1.8|
|[createWorkbookPostProcessAction](/javascript/api/excel/excel.createworkbookpostprocessaction)|_属性_> fakeFileId|传输到客户端，例如，对于 TableSelectionChangedEvent worksheetId 的其他数据。|1.8|
|[createWorkbookPostProcessAction](/javascript/api/excel/excel.createworkbookpostprocessaction)|_属性_> fileBase64|传输到客户端，例如，对于 TableSelectionChangedEvent worksheetId 的其他数据。|1.8|
|[createWorkbookPostProcessAction](/javascript/api/excel/excel.createworkbookpostprocessaction)|_关系_> actionType|传输到客户端，例如，对于 TableSelectionChangedEvent worksheetId 的其他数据。|1.8|
|[customDataValidation](/javascript/api/excel/excel.customdatavalidation)|_属性_> 公式| 自定义数据验证公式。 这将创建特殊输入的规则，例如阻止重复或限制单元格区域中的总数。|1.8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_属性_ > id|DataPivotHierarchy 的 id。 只读。|1.8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_属性_ > name|DataPivotHierarchy 的名称。|1.8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_属性_ > numberFormat|DataPivotHierarchy 数字格式。|1.8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_属性_> 位置|DataPivotHierarchy 位置。|1.8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_关系_> 字段|返回与 DataPivotHierarchy 关联的数据透视。 只读。|1.8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_关系_> showAs|确定数据是否应为特定的汇总计算或不显示。|1.8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_关系_> summarizeBy|确定是否显示 DataPivotHierarchy 的所有项目。|1.8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_方法_ > [setToDefault()](/javascript/api/excel/excel.datapivothierarchy#settodefault)|DataPivotHierarchy 重置为其默认值。|1.8|
|[dataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|_属性_ > items|DataPivotHierarchy 对象的集合。 只读。|1.8|
|[dataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|_方法_ > [add(pivotHierarchy: PivotHierarchy)](/javascript/api/excel/excel.datapivothierarchycollection)|向当前轴 PivotHierarchy。|1.8|
|[dataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|_方法_  >  [getCount()](/javascript/api/excel/excel.datapivothierarchycollection)|获取集合中的数据透视表层次结构的数量。|1.8|
|[dataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|_方法_ > [getItem(name: 字符串)](/javascript/api/excel/excel.datapivothierarchycollection)|获取 DataPivotHierarchy 按其名称或 id。|1.8|
|[dataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|_方法_  >  [getItemOrNullObject(name: string)](/javascript/api/excel/excel.datapivothierarchycollection)|按名称获取 DataPivotHierarchy。 如果 DataPivotHierarchy 不存在，将返回 null 对象。|1.8|
|[dataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|_方法_ > [remove(DataPivotHierarchy: DataPivotHierarchy)](/javascript/api/excel/excel.datapivothierarchycollection)|删除当前轴 PivotHierarchy。|1.8|
|[验证说明数据验证](/javascript/api/excel/excel.datavalidation)|_属性_> ignoreBlanks|忽略空值： 无数据验证将空白单元格上执行，其默认值为 true。|1.8|
|[验证说明数据验证](/javascript/api/excel/excel.datavalidation)|_属性_> 有效|代表所有单元格的值是否有效根据数据验证规则。 只读。|1.8|
|[验证说明数据验证](/javascript/api/excel/excel.datavalidation)|_关系_> errorAlert|错误警报，如果用户输入无效数据。|1.8|
|[验证说明数据验证](/javascript/api/excel/excel.datavalidation)|_关系_> 提示|当用户选择单元格的提示。|1.8|
|[验证说明数据验证](/javascript/api/excel/excel.datavalidation)|_关系_> 规则|包含不同类型的数据验证条件的数据有效性规则。|1.8|
|[验证说明数据验证](/javascript/api/excel/excel.datavalidation)|_关系_ > type|键入的数据验证，有关详细信息，请参阅[Excel.DataValidationType](/javascript/api/excel/excel.datavalidationtype) 。 只读。|1.8|
|[验证说明数据验证](/javascript/api/excel/excel.datavalidation)|_方法_ > [clear()](/javascript/api/excel/excel.datavalidation)|清除当前范围中的数据有效性。|1.8|
|[dataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)|_属性_> 邮件|代表错误警告消息。|1.8|
|[dataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)|_属性_> showAlert|确定是否显示错误的警报对话框或不时用户输入无效数据。 默认值为 true。|1.8|
|[dataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)|_属性_ > title|代表错误警报对话框标题。|1.8|
|[dataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)|_关系_> 样式|代表数据验证通知类型，有关详细信息，请参阅[Excel.DataValidationAlertStyle](/javascript/api/excel/excel.datavalidationalertstyle) 。|1.8|
|[dataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt)|_属性_> 邮件|表示的消息的提示。|1.8|
|[dataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt)|_属性_> showPrompt|确定显示提示，当用户选择单元格的数据有效性。|1.8|
|[dataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt)|_属性_ > title|代表提示符处的标题。|1.8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule)|_关系_> 自定义|自定义数据验证条件。|1.8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule)|_关系_> 日期|日期数据验证条件。|1.8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule)|_关系_> 十进制|十进制数据验证条件。|1.8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule)|_关系_ > list|列表数据验证条件。|1.8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule)|_关系_> textLength|TextLength 数据验证条件。|1.8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule)|_关系_> 时间|时间数据验证条件。|1.8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule)|_关系_> wholeNumber|WholeNumber 数据验证条件。|1.8|
|[dateTimeDataValidation](/javascript/api/excel/excel.datetimedatavalidation)|_属性_> formula1|获取或设置 Formula1，即最小值或根据运算符的值。|1.8|
|[dateTimeDataValidation](/javascript/api/excel/excel.datetimedatavalidation)|_属性_> formula2|获取或设置 Formula2，即最大值或根据运算符的值。|1.8|
|[dateTimeDataValidation](/javascript/api/excel/excel.datetimedatavalidation)|_关系_> 运算符|要用于验证的数据的运算符。|1.8|
|[enableEventsPostProcessAction](/javascript/api/excel/excel.enableeventspostprocessaction)|_属性_> isEnableEvents {|传输到客户端，例如，对于 TableSelectionChangedEvent worksheetId 的其他数据。|1.8|
|[enableEventsPostProcessAction](/javascript/api/excel/excel.enableeventspostprocessaction)|_关系_> actionType|传输到客户端，例如，对于 TableSelectionChangedEvent worksheetId 的其他数据。|1.8|
|[enableEventsPostProcessAction](/javascript/api/excel/excel.enableeventspostprocessaction)|_关系_> controlId|传输到客户端，例如，对于 TableSelectionChangedEvent worksheetId 的其他数据。|1.8|
|[filterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|_属性_> enableMultipleFilterItems|确定是否允许多个筛选项目。|1.8|
|[filterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|_属性_ > id|FilterPivotHierarchy 的 id。 只读。|1.8|
|[filterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|_属性_ > name|FilterPivotHierarchy 的名称。|1.8|
|[filterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|_属性_> 位置|FilterPivotHierarchy 位置。|1.8|
|[filterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|_关系_ > fields|返回与 FilterPivotHierarchy 关联的数据透视。 只读。|1.8|
|[filterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|_方法_ > [setToDefault()](/javascript/api/excel/excel.filterpivothierarchy)|FilterPivotHierarchy 重置为其默认值。|1.8|
|[filterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|_属性_ > items|FilterPivotHierarchy 对象的集合。 只读。|1.8|
|[filterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|_方法_ > [add(pivotHierarchy: PivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchycollection)|向当前轴 PivotHierarchy。 如果存在行、 列或筛选轴上的其他位置层次结构，它将从该位置。|1.8|
|[filterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|_方法_  >  [getCount()](/javascript/api/excel/excel.filterpivothierarchycollection)|获取集合中的数据透视表层次结构的数量。|1.8|
|[filterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|_方法_ > [getItem(name: 字符串)](/javascript/api/excel/excel.filterpivothierarchycollection)|获取 FilterPivotHierarchy 按其名称或 id。|1.8|
|[filterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|_方法_  >  [getItemOrNullObject(name: string)](/javascript/api/excel/excel.filterpivothierarchycollection)|按名称获取 FilterPivotHierarchy。 如果 FilterPivotHierarchy 不存在，将返回 null 对象。|1.8|
|[filterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|_方法_ > [remove(filterPivotHierarchy: FilterPivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchycollection)|删除当前轴 PivotHierarchy。|1.8|
|[listDataValidation](/javascript/api/excel/excel.listdatavalidation)|_属性_> inCellDropDown|显示或不下拉列表中单元格，则默认为 true。|1.8|
|[listDataValidation](/javascript/api/excel/excel.listdatavalidation)|_属性_> 源|源的数据有效性的列表|1.8|
|[openWorkbookPostProcessAction](/javascript/api/excel/excel.openworkbookpostprocessaction)|_属性_> fakeFileId|传输到客户端，例如，对于 TableSelectionChangedEvent worksheetId 的其他数据。|1.8|
|[openWorkbookPostProcessAction](/javascript/api/excel/excel.openworkbookpostprocessaction)|_关系_> actionType|传输到客户端，例如，对于 TableSelectionChangedEvent worksheetId 的其他数据。|1.8|
|[pivotField](/javascript/api/excel/excel.pivotfield)|_属性_ > id|透视字段的 id。 只读。|1.8|
|[pivotField](/javascript/api/excel/excel.pivotfield)|_属性_ > name|透视字段的名称。|1.8|
|[pivotField](/javascript/api/excel/excel.pivotfield)|_属性_> showAllItems|确定是否显示透视字段的所有项。|1.8|
|[pivotField](/javascript/api/excel/excel.pivotfield)|_关系_> 项目|返回与 PivotField 关联的数据透视。 只读。|1.8|
|[pivotField](/javascript/api/excel/excel.pivotfield)|_关系_> 分类汇总|分类汇总的透视字段。|1.8|
|[pivotField](/javascript/api/excel/excel.pivotfield)|_方法_ > [sortByLabels(sortby: SortBy)](/javascript/api/excel/excel.pivotfield)|对透视字段进行排序。 如果指定 DataPivotHierarchy，则排序将应用基于，如果没有排序将基于本身的 PivotField。|1.8|
|[pivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection)|_属性_ > items|PivotField 对象的集合。 只读。|1.8|
|[pivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection)|_方法_  >  [getCount()](/javascript/api/excel/excel.pivotfieldcollection)|获取集合中的数据透视表层次结构的数量。|1.8|
|[pivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection)|_方法_ > [getItem(name: 字符串)](/javascript/api/excel/excel.pivotfieldcollection)|获取 PivotHierarchy 按其名称或 id。|1.8|
|[pivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection)|_方法_  >  [getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivotfieldcollection)|按名称获取 PivotHierarchy。 如果 PivotHierarchy 不存在，将返回 null 对象。|1.8|
|[pivotHierarchy](/javascript/api/excel/excel.pivothierarchy)|_属性_ > id|PivotHierarchy 的 id。 只读。|1.8|
|[pivotHierarchy](/javascript/api/excel/excel.pivothierarchy)|_属性_ > name|PivotHierarchy 的名称。|1.8|
|[pivotHierarchy](/javascript/api/excel/excel.pivothierarchy)|_关系_ > fields|返回与 PivotHierarchy 关联的数据透视。 只读。|1.8|
|[pivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection)|_属性_ > items|PivotHierarchy 对象的集合。 只读。|1.8|
|[pivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection)|_方法_  >  [getCount()](/javascript/api/excel/excel.pivothierarchycollection)|获取集合中的数据透视表层次结构的数量。|1.8|
|[pivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection)|_方法_ > [getItem(name: 字符串)](/javascript/api/excel/excel.pivothierarchycollection)|获取 PivotHierarchy 按其名称或 id。|1.8|
|[pivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection)|_方法_  >  [getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivothierarchycollection)|按名称获取 PivotHierarchy。 如果 PivotHierarchy 不存在，将返回 null 对象。|1.8|
|[pivotItem](/javascript/api/excel/excel.pivotitem)|_属性_ > id|PivotItem 的 id。 只读。|1.8|
|[pivotItem](/javascript/api/excel/excel.pivotitem)|_属性_> isExpanded|确定项目是否已展开以显示子项目，或如果折叠和已被隐藏子项目。|1.8|
|[pivotItem](/javascript/api/excel/excel.pivotitem)|_属性_ > name|PivotItem 的名称。|1.8|
|[pivotItem](/javascript/api/excel/excel.pivotitem)|_属性_> 可见|确定 PivotItem 是否可见。|1.8|
|[pivotItemCollection](/javascript/api/excel/excel.pivotitemcollection)|_属性_ > items|PivotItem 对象的集合。 只读。|1.8|
|[pivotItemCollection](/javascript/api/excel/excel.pivotitemcollection)|_方法_  >  [getCount()](/javascript/api/excel/excel.pivotitemcollection)|获取集合中的数据透视表层次结构的数量。|1.8|
|[pivotItemCollection](/javascript/api/excel/excel.pivotitemcollection)|_方法_ > [getItem(name: 字符串)](/javascript/api/excel/excel.pivotitemcollection)|获取 PivotHierarchy 按其名称或 id。|1.8|
|[pivotItemCollection](/javascript/api/excel/excel.pivotitemcollection)|_方法_  >  [getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivotitemcollection)|按名称获取 PivotHierarchy。 如果 PivotHierarchy 不存在，将返回 null 对象。|1.8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_属性_> showColumnGrandTotals|如果数据透视表报表显示总计总计列，则为 true。|1.8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_属性_> showRowGrandTotals|如果数据透视表报表显示总计行的总数。|1.8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_属性_> subtotalLocation|此属性指示对数据透视表的所有字段的 SubtotalLocationType。 如果字段具有不同的状态，这将为空。 可能的值为： AtTop、 AtBottom。|1.8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_关系_> layoutType|此属性指示对数据透视表的所有字段的 PivotLayoutType。 如果字段具有不同的状态，这将为空。|1.8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_方法_ > [getColumnLabelRange()](/javascript/api/excel/excel.pivotlayout)|返回数据透视表的列标签所在的区域。|1.8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_方法_ > [getDataBodyRange()](/javascript/api/excel/excel.pivotlayout)|返回数据透视表的数据值所在的区域。|1.8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout.md)|_方法_ > [getFilterAxisRange()](/javascript/api/excel/excel.pivotlayout)|返回数据透视表的筛选器区域的范围。|1.8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_方法_ > [getRange()](/javascript/api/excel/excel.pivotlayout)|返回位于数据透视表，排除筛选器区域的范围。|1.8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_方法_ > [getRowLabelRange()](/javascript/api/excel/excel.pivotlayout)|返回数据透视表的行标签所在的区域。|1.8|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_关系_> columnHierarchies|数据透视表列数据透视表层次结构中。 只读。|1.8|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_关系_> dataHierarchies|数据透视表的数据透视层次结构中。 只读。|1.8|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_关系_> filterHierarchies|数据透视表筛选器数据透视表层次结构中。 只读。|1.8|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_关系_> 层次结构|数据透视表的数据透视表层次结构中。 只读。|1.8|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_关系_> 布局|用于说明布局和数据透视表的可视结构 PivotLayout。 只读。|1.8|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_关系_> rowHierarchies|数据透视表行数据透视表层次结构中。 只读。|1.8|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_方法_ > [delete()](/javascript/api/excel/excel.pivottable)|删除数据透视表。|1.8|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|_方法_ > [add(name: string, source: object, destination: object)](/javascript/api/excel/excel.pivottablecollection)|添加数据透视表基于指定的源数据并将其插入目标区域左上角单元格。|1.8|
|[range](/javascript/api/excel/excel.range)|_关系_> 验证说明数据验证|返回的数据验证对象。 只读。|1.8|
|[rowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|_属性_ > id|RowColumnPivotHierarchy 的 id。 只读。|1.8|
|[rowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|_属性_ > name|RowColumnPivotHierarchy 的名称。|1.8|
|[rowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|_属性_> 位置|RowColumnPivotHierarchy 位置。|1.8|
|[rowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|_关系_ > fields|返回与 RowColumnPivotHierarchy 关联的数据透视。 只读。|1.8|
|[rowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|_方法_ > [setToDefault()](/javascript/api/excel/excel.rowcolumnpivothierarchy)|RowColumnPivotHierarchy 重置为其默认值。|1.8|
|[rowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|_属性_ > items|RowColumnPivotHierarchy 对象的集合。 只读。|1.8|
|[rowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|_方法_ > [add(pivotHierarchy: PivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|向当前轴 PivotHierarchy。 层次结构上是否存在其他地方行中，列中，|1.8|
|[rowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|_方法_  >  [getCount()](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|获取集合中的数据透视表层次结构的数量。|1.8|
|[rowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|_方法_ > [getItem(name: 字符串)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|获取 RowColumnPivotHierarchy 按其名称或 id。|1.8|
|[rowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|_方法_  >  [getItemOrNullObject(name: string)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|按名称获取 RowColumnPivotHierarchy。 如果 RowColumnPivotHierarchy 不存在，将返回 null 对象。|1.8|
|[rowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|_方法_ > [remove(rowColumnPivotHierarchy: RowColumnPivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|删除当前轴 PivotHierarchy。|1.8|
|[运行时](/javascript/api/excel/excel.runtime)|_属性_> enableEvents|切换 JavaScript 中的当前任务窗格或内容加载项的事件。|1.8|
|[showAsRule](/javascript/api/excel/excel.showasrule)|_关系_> baseField|基本透视字段中，如果适用，则基于 ShowAs 计算，基于 else null ShowAsCalculation 类型。|1.8|
|[showAsRule](/javascript/api/excel/excel.showasrule)|_关系_> baseItem|基本的项，如果适用，基于 ShowAs 计算基于 else null ShowAsCalculation 类型。|1.8|
|[showAsRule](/javascript/api/excel/excel.showasrule)|_关系_> 计算|要用于数据透视字段 ShowAs 计算。|1.8|
|[style](/javascript/api/excel/excel.style)|_属性_> 自动|指示是否文本会自动缩进单元格中的文本对齐方式设置为等距分布时。|1.8|
|[style](/javascript/api/excel/excel.style)|_属性_> textOrientation|样式的文本方向。|1.8|
|[分类汇总](/javascript/api/excel/excel.subtotals)|_属性_> 自动|如果自动设置为 true，则所有其他值将被忽略时设置分类汇总。|1.8|
|[分类汇总](/javascript/api/excel/excel.subtotals)|_属性_> 平均| |1.8|
|[分类汇总](/javascript/api/excel/excel.subtotals)|_属性_> 计数| |1.8|
|[分类汇总](/javascript/api/excel/excel.subtotals)|_属性_> countNumbers| |1.8|
|[分类汇总](/javascript/api/excel/excel.subtotals)|_属性_> 最大| |1.8|
|[分类汇总](/javascript/api/excel/excel.subtotals)|_属性_> 最小值| |1.8|
|[分类汇总](/javascript/api/excel/excel.subtotals)|_属性_> 产品| |1.8|
|[分类汇总](/javascript/api/excel/excel.subtotals)|_属性_> standardDeviation| |1.8|
|[分类汇总](/javascript/api/excel/excel.subtotals)|_属性_> standardDeviationP| |1.8|
|[分类汇总](/javascript/api/excel/excel.subtotals)|_属性_> sum| |1.8|
|[分类汇总](/javascript/api/excel/excel.subtotals)|_属性_> 差异| |1.8|
|[分类汇总](/javascript/api/excel/excel.subtotals)|_属性_> varianceP| |1.8|
|[table](/javascript/api/excel/excel.table)|_属性_> legacyId|返回数字的 id。只读的。|1.8|
|[workbook](/javascript/api/excel/excel.workbook)|_属性_> readOnly|如果在只读模式中打开工作簿，则为 true。 只读。|1.8|
|[workbookCreated](/javascript/api/excel/excel.workbookcreated)|_属性_ > id|返回一个值，用于唯一标识 WorkbookCreated 对象。 只读。|1.8|
|[workbookCreated](/javascript/api/excel/excel.workbookcreated)|_方法_ > [open()](/javascript/api/excel/excel.workbookcreated)|打开工作簿。|1.8|
|[worksheet](/javascript/api/excel/excel.worksheet)|_属性_> showGridlines|获取或设置工作表的网格线标志。|1.8|
|[worksheet](/javascript/api/excel/excel.worksheet)|_属性_> showHeadings|获取或设置工作表的标题的标志。|1.8|
|[worksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|_属性_> 类型|获取该事件的类型。|1.8|
|[worksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|_属性_> worksheetId|获取的计算工作表的 id。|1.8|

## <a name="whats-new-in-excel-javascript-api-17"></a>What's new in Excel JavaScript API 1.7

Excel 的 JavaScript API 要求集 1.7 功能包括图表、 事件、 工作表、 区域、 文档属性，为项目、 保护选项和样式的 Api。

### <a name="customize-charts"></a>自定义图表

与新图表 Api，您可以创建其他图表类型、 向图表中添加一个数据系列，设置图表标题、 添加坐标轴标题、 添加显示单位、 添加与移动平均趋势线、 更改为线性趋势线和更多内容。 以下是一些示例：

* 图表轴-获取、 设置、 格式和图表中删除轴单位、 标签和标题。
* 图表系列-添加、 设置和删除图表中的一系列。  更改系列标记、 绘图订单和大小。
* 图表趋势线-添加、 获取和设置图表中的趋势线的格式。
* 图表图例-图表中的图例字体格式。
* 图表点-设置图表点颜色。
* 图表标题的子字符串-获取和设置图表标题子字符串。
* 图表类型-创建更多的图表类型的选项。

### <a name="events"></a>事件

Api 提供允许加载项以自动运行指定的函数时特定事件的事件处理程序的各种 Excel 事件发生。 可以将函数设计为执行方案所需的任何操作。 当前可用的事件的列表，请参阅[使用使用 Excel JavaScript API 的事件](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-events)。

### <a name="customize-the-appearance-of-worksheets-and-ranges"></a>自定义工作表和区域的外观

使用新的 Api，您可以自定义工作表以多种方式的外观：

* 冻结窗格保留特定的行或列可见滚动工作表中时。 例如，如果在工作表中的第一行包含标头，您可能冻结该行，以便向工作表下滚动，仍可见列标题。
* 修改工作表标签颜色。
* 添加表的标题。


您可以自定义区域以多种方式的外观：

* 设置以确保范围的单元格样式确保范围中的所有单元格具有一致的格式。 单元格样式是一组定义的格式特征，如字体和字体大小、 数字格式、 单元格边框和单元格底纹。 可以使用任何 Excel 的内置单元格样式或创建您自己的自定义的单元格样式。
* 设置范围内的文本方向。
* 添加或修改区域链接到工作簿中的另一个位置或外部位置上的超链接。

### <a name="manage-document-properties"></a>管理文档属性

使用 Api 中的文档属性，您可以访问内置文档属性和还创建和管理自定义文档属性，用于存储工作簿和驱动器的工作流和业务逻辑的状态。

### <a name="copy-worksheets"></a>复制工作表

使用工作表复制 Api，可以将从一个工作表的数据和格式复制到同一工作簿中的新工作表并减少所需的数据传输的量。

### <a name="handle-ranges-with-ease"></a>轻松地处理范围

使用各种范围 Api，您可以执行诸如 get 周围的区域，获取调整大小的范围和详细信息。 这些 Api 应使任务，如范围操作和解决更加有效。

此外：

* 工作簿和工作表的保护选项-使用这些 Api 来保护工作表和工作簿结构中的数据。
* 更新已命名的项目-使用此 API 更新已命名的项目。
* 获取活动单元格-使用此 API 获取工作簿的活动单元格。

|对象| 最近更新| 说明|要求集|
|:----|:----|:----|:----|
|[chart](/javascript/api/excel/excel.chart)|_属性_> chartType|代表图表的类型。 可能的值为： ColumnClustered、 ColumnStacked、 ColumnStacked100、 BarClustered、 BarStacked、 BarStacked100、 LineStacked、 LineStacked100、 LineMarkers、 LineMarkersStacked、 LineMarkersStacked100、 PieOfPie，等...|1.7|
|[chart](/javascript/api/excel/excel.chart)|_属性_ > id|图表的唯一 id。 只读。|1.7|
|[chart](/javascript/api/excel/excel.chart)|_属性_> showAllFieldButtons|代表是否在数据透视图上显示所有字段按钮。|1.7|
|[chartAreaFormat](/javascript/api/excel/excel.chartareaformat)|_关系_> 边框|代表图表区域，其中包括颜色、 linestyle 和权重的边框格式。 只读。|1.7|
|[chartAxes](/javascript/api/excel/excel.chartaxes)|_方法_> getItem (类型： string、 组： 字符串)|返回由类型和组标识的特定轴。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_属性_> axisBetweenCategories|表示是否数值轴与分类坐标轴相交于分类之间。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_属性_> axisGroup|代表指定坐标轴组。 只读。 可能的值为： 主、 辅助。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_属性_> categoryType|返回或设置分类轴类型。 可能的值为： 自动，TextAxis，DateAxis。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_属性_> 交叉|代表指定的坐标轴与其他坐标轴相交。 可能的值为： 自动、 最大值、 最小、 自定义。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_属性_> crossesAt|代表指定的坐标轴与其他坐标轴交叉点。 只读。 设置为此属性应使用 SetCrossesAt(double) 方法。 只读。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_属性_> customDisplayUnit|代表自定义轴显示单位的值。 只读。 若要设置该属性，请使用 SetCustomDisplayUnit(double) 方法。 只读。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_属性_> displayUnit|代表轴显示单位。 可能的值： None、 数百、 千位、 TenThousands、 HundredThousands、 数百万、 TenMillions、 HundredMillions、 十亿、 Trillions、 自定义。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_属性_ > height|表示高度，以磅为单位的图表坐标轴。 如果坐标轴的不可见，则为 null。 只读。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_属性_> 左侧|表示从左侧的图表区的轴的左边缘的距离，以磅为单位。 如果坐标轴的不可见，则为 null。 只读。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_属性_> logBase|代表对数的基时使用对数刻度。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_属性_> reversePlotOrder|表示是否 Microsoft Excel 数据点绘制数据从最后一到第一个。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_属性_> scaleType|代表数值轴的刻度类型。 可能的值为： 线性、 对数。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_属性_> showDisplayUnitLabel|代表轴显示单位标签是否可见。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_属性_> tickLabelSpacing|表示类别或刻度线标签之间的系列的数量。 可以是介于 1 至 31999 或自动设置为空字符串。 返回的值始终是一个数字。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_属性_> tickMarkSpacing|表示类别或刻度线之间的系列的数量。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_属性_> 顶部|代表中的图表区域顶部的轴的上边缘的距离，以磅为单位。 如果坐标轴的不可见，则为 null。 只读。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_属性_> 类型|代表坐标轴类型。 只读。 可能的值为： 无效、 类别、 值、 系列。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_属性_> 可见|一个布尔值代表坐标轴的可见性。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_属性_ > width|表示的宽度，以磅为单位的图表坐标轴。 如果坐标轴的不可见，则为 null。 只读。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_关系_> baseTimeUnit|返回或设置指定分类轴的基本单位。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_关系_> majorTickMark|表示指定坐标轴的主要刻度线标志的类型。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_关系_> majorTimeUnitScale|返回或设置分类轴的主要刻度单位扩展值，当 CategoryType 属性设置为时间刻度。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_关系_> minorTickMark|表示指定坐标轴的次要刻度线的类型。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_关系_> minorTimeUnitScale|返回或设置分类轴的次要刻度单位扩展值，当 CategoryType 属性设置为时间刻度。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_关系_> tickLabelPosition|代表指定坐标轴上刻度线标签的位置。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_方法_> setCategoryNames(sourceData: Range)|设置指定坐标轴的所有类别名称。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_方法_> setCrossesAt(value: double)|设置指定的坐标轴与其他坐标轴交叉点。|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_方法_> setCustomDisplayUnit(value: double)|将轴显示单位设置为自定义值。|1.7|
|[chartBorder](/javascript/api/excel/excel.chartborder)|_属性_ > color|代表图表中的边框的颜色的 HTML 颜色代码。|1.7|
|[chartBorder](/javascript/api/excel/excel.chartborder)|_属性_> 权重|表示边框，以磅为单位的权重。|1.7|
|[chartBorder](/javascript/api/excel/excel.chartborder)|_关系_> lineStyle|代表边框的线型。|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_属性_> 位置|表示数据标签的位置的 DataLabelPosition 值。可能的值是：None、Center、InsideEnd、InsideBase、OutsideEnd、Left、Right、Top、Bottom、BestFit、Callout。|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_属性_> 分隔符|字符串，表示用于图表上的数据标签的分隔符。|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_属性_> showBubbleSize|表示数据标签气泡大小是否可见的布尔值。|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_属性_> showCategoryName|表示数据标签类别名称是否可见的布尔值。|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_属性_> showLegendKey|表示数据标签图例标示是否可见的布尔值。|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_属性_> showPercentage|表示数据标签百分比是否可见的布尔值。|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_属性_> showSeriesName|表示数据标签系列名称是否可见的布尔值。|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_属性_> showValue|表示数据标签值是否可见的布尔值。|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_属性_ > height|代表图表上图例的高度。|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_属性_> 左侧|代表图表图例的左侧。|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_属性_> showShadow|代表如果图例在图表上有阴影。|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_属性_> 顶部|代表图表图例的顶部。|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_属性_ > width|代表图表上图例的宽度。|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_关系_> legendEntries|代表在图例中的 legendEntries 集合。 只读。|1.7|
|[chartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|_属性_> 可见|代表图表图例项的可见。|1.7|
|[chartLegendEntryCollection](/javascript/api/excel/excel.chartlegendentrycollection)|_属性_ > items|ChartLegendEntry 对象的集合。 只读。|1.7|
|[chartLegendEntryCollection](/javascript/api/excel/excel.chartlegendentrycollection)|_方法_> getCount()|返回集合中的 legendEntry 数。|1.7|
|[chartLegendEntryCollection](/javascript/api/excel/excel.chartlegendentrycollection)|_方法_> getItemAt(index: number)|返回给定索引处 legendEntry。|1.7|
|[chartPoint](/javascript/api/excel/excel.chartpoint)|_属性_> hasDataLabel|表示是否数据点具有 datalabel。 不适用于曲面图。|1.7|
|[chartPoint](/javascript/api/excel/excel.chartpoint)|_属性_> markerBackgroundColor|HTML 颜色代码表示的数据标记背景色的点。 例如， #FF0000 表示红色。|1.7|
|[chartPoint](/javascript/api/excel/excel.chartpoint)|_属性_> markerForegroundColor|HTML 颜色代码表示形式的标记前景色的数据点。 例如， #FF0000 表示红色。|1.7|
|[chartPoint](/javascript/api/excel/excel.chartpoint)|_属性_> markerSize|表示数据点的标记的大小。|1.7|
|[chartPoint](/javascript/api/excel/excel.chartpoint)|_属性_> markerStyle|代表图表数据点的标记样式。 可能的值为： 无效、 自动、 无、 方形、 菱形、 三角形，X、 星形、 点、 划线、 Circle，以及，图片。|1.7|
|[chartPoint](/javascript/api/excel/excel.chartpoint)|_关系_> dataLabel|返回图表点的数据标签。 只读。|1.7|
|[chartPointFormat](/javascript/api/excel/excel.chartpointformat)|_关系_> 边框|代表图表数据点，其中包括颜色、 样式和权重的信息的边框格式。 只读。|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_属性_> chartType|表示一系列的图表类型。 可能的值为： ColumnClustered、 ColumnStacked、 ColumnStacked100、 BarClustered、 BarStacked、 BarStacked100、 LineStacked、 LineStacked100、 LineMarkers、 LineMarkersStacked、 LineMarkersStacked100、 PieOfPie，等...|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_属性_> doughnutHoleSize|代表图表系列的圆环图内径大小。  仅圆环图和 doughnutExploded 图表上有效。|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_属性_> 筛选|布尔值，表示如果或不筛选数据系列。 不适用于曲面图。|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_属性_> gapWidth|代表图表系列的间距。  唯一有效上栏和柱形图，以及|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_属性_> hasDataLabels|布尔值，表示如果系列或不具有数据标签。|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_属性_> markerBackgroundColor|代表图表系列的标记背景色。|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_属性_> markerForegroundColor|代表图表系列的标记前景色。|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_属性_> markerSize|代表图表系列的标记大小。|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_属性_> markerStyle|代表图表系列的数据标记样式。 可能的值为： 无效、 自动、 无、 方形、 菱形、 三角形，X、 星形、 点、 划线、 Circle，以及，图片。|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_属性_> plotOrder|代表图表组中的图表系列的绘制顺序。|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_属性_> showShadow|布尔值，表示如果系列有阴影，或不。|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_属性_> 顺利|布尔值，表示系列是否平滑。 仅对折线图、 散点图。|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_关系_> dataLabels|代表系列中的所有数据标签的集合。 只读。|ApiSet.InProgressFeatures.ChartingAPI|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_关系_> 趋势线|代表系列中的趋势线的集合。 只读。|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_方法_> delete()|删除图表数据系列。|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_方法_> setBubbleSizes(sourceData: Range)|设置图表序列的气泡大小。 仅适用于气泡图。|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_方法_> setValues(sourceData: Range)|设置图表序列的值。 对于散点图图表，它表示的 Y 轴值。|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_方法_> setXAxisValues(sourceData: Range)|设置的 X 轴图表系列的值。 仅适用于散点图。|1.7|
|[chartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|_方法_> 添加 (名称： string、 索引： 数)|向集合添加一个新系列。|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_属性_ > height|返回高度，以磅为单位的图表标题。 只读。 如果图表标题的不可见，则为 null。 只读。|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_属性_> horizontalAlignment|代表图表标题的水平对齐方式。 可能的值为： 中心，左、 两端对齐，分布式，右。|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_属性_> 左侧|代表中图表区的左边缘到图表标题的左边缘的距离，以磅为单位。 如果图表标题的不可见，则为 null。|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_属性_> 位置|代表图表标题的位置。 可能的值为： 顶部、 自动、 底部、 右、 左侧。|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_属性_> showShadow|代表一个布尔值，确定图表标题是否具有阴影。|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_属性_> textOrientation|代表图表标题的文本方向。 值应为整数从-90 至 90 或 180 纵向文本。|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_属性_> 顶部|代表中的图表区域顶部的图表标题的上边缘的距离，以磅为单位。 如果图表标题的不可见，则为 null。|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_属性_> verticalAlignment|代表图表标题的垂直对齐方式。 可能的值为： 中心、 底部、 Top、 两端对齐、 分布式。|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_属性_ > width|返回宽度，以磅为单位的图表标题。 只读。 如果图表标题的不可见，则为 null。 只读。|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_方法_> setFormula(formula: string)|设置一个 string 值，它代表图表标题使用 A1 样式表示法的公式。|1.7|
|[chartTitleFormat](/javascript/api/excel/excel.charttitleformat)|_关系_> 边框|代表图表标题，其中包括颜色、 linestyle 和权重的边框格式。 只读。|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_属性_> 向后|代表趋势线向后延伸的周期数。|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_属性_> displayEquation|如果在图表上显示趋势线公式，则为 true。|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_属性_> displayRSquared|如果 R 平方趋势线在图表上显示，则为 true。|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_属性_> 转接|代表趋势线向前延伸的周期数。|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_属性_> 截距|代表趋势线的截距值。 可以设置为数字值或空字符串 （用于自动值）。 返回的值始终是一个数字。|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_属性_> movingAveragePeriod|代表图表中趋势线，仅趋势线的周期移动平均类型。|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_属性_ > name|代表趋势线的名称。 可以将设置为一个字符串值，或者可以设置为 null 值表示自动值。 返回的值始终为字符串|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_属性_> polynomialOrder|代表图表中趋势线，仅趋势线的顺序与多项式类型。|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_属性_> 类型|代表图表中趋势线的类型。 可能的值为： 线性、 指数、 对数，移动平均多项式、 电源。|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_关系_> 格式|代表图表中趋势线的格式。 只读。|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_方法_> delete()|删除 trendline 对象。|1.7|
|[chartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|_属性_ > items|ChartTrendline 对象的集合。 只读。|1.7|
|[chartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|_方法_> add(type: string)|向趋势线集合添加新的趋势线。|1.7|
|[chartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|_方法_> getCount()|返回集合中的趋势线数目。|1.7|
|[chartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|_方法_> getItem(index: number)|通过索引，这是数组中的项的插入顺序来获取 trendline 对象。|1.7|
|[chartTrendlineFormat](/javascript/api/excel/excel.charttrendlineformat)|_关系_> 行|表示图表线条格式。只读。|1.7|
|[customProperty](/javascript/api/excel/excel.customproperty)|_属性_ > key|获取 customProperty 的键。只读。只读。|1.7|
|[customProperty](/javascript/api/excel/excel.customproperty)|_属性_> 类型|获取自定义属性的值类型。 只读。 只读。 可能的值为： 编号、 Boolean、 日期、 字符串、 浮点数。|1.7|
|[customProperty](/javascript/api/excel/excel.customproperty)|_属性_ > value|获取或设置 customProperty 的值。|1.7|
|[customProperty](/javascript/api/excel/excel.customproperty)|_方法_> delete()|删除 customProperty。|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_属性_ > items|一组 CustomProperty 对象。只读。|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_方法_> 添加 (密钥： 字符串，值： 对象)|新建自定义属性或设置现有自定义属性。|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_方法_> deleteAll()|删除此集合中的所有 customProperty。|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_方法_> getCount()|获取 customProperty 的计数。|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_方法_> getItem(key: string)|按键获取 custom property 对象（不区分大小写）。当不存在自定义属性时引发。|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_方法_> getItemOrNullObject(key: string)|按键获取 custom property 对象（不区分大小写）。如果不存在自定义属性，则返回 null 对象。|1.7|
|[dataConnectionCollection](/javascript/api/excel/excel.dataconnectioncollection)|_属性_ > items|DataConnection 对象的集合。 只读。|1.7|
|[dataConnectionCollection](/javascript/api/excel/excel.dataconnectioncollection)|_方法_> refreshAll()|刷新集合中的所有数据连接。|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_属性_ > author|获取或设置工作簿的作者。|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_属性_ > category|获取或设置工作簿的类别。|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_属性_ > comments|获取或设置工作簿的注释。|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_属性_ > company|获取或设置工作簿的公司。|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_属性_ > keywords|获取或设置工作簿的关键字。|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_属性_ > lastAuthor|获取工作簿的上一个作者。 只读。 只读。|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_属性_ > manager|获取或设置工作簿的经理。|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_属性_ > revisionNumber|获取工作簿的修订号。 只读。|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_属性_ > subject|获取或设置工作簿的主题。|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_属性_ > title|获取或设置工作簿的标题。|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_关系_ > creationDate|获取工作簿的创建日期。 只读。 只读。|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_关系_> 自定义|获取工作簿的自定义属性的集合。 只读。 只读。|1.7|
|[namedItem](/javascript/api/excel/excel.nameditem)|_属性_> 公式|获取或设置命名项的公式。  公式始终开头 = 登录。|1.7|
|[namedItem](/javascript/api/excel/excel.nameditem)|_关系_> arrayValues|返回一个包含值和类型的已命名项目对象。 只读。|1.7|
|[namedItemArrayValues](/javascript/api/excel/excel.nameditemarrayvalues)|_属性_> 类型|代表已命名的项目数组中每个项的类型只读的。 可能的值为： 未知、 空、 String、 Integer、 Double、 Boolean、 错误。|1.7|
|[namedItemArrayValues](/javascript/api/excel/excel.nameditemarrayvalues)|_属性_ > values|表示已命名的项目数组中每个项的值。 只读。|1.7|
|[range](/javascript/api/excel/excel.range)|_属性_> isEntireColumn|代表当前范围是否整列。 只读。|1.7|
|[range](/javascript/api/excel/excel.range)|_属性_> isEntireRow|代表当前范围是否整行。 只读。|1.7|
|[range](/javascript/api/excel/excel.range)|_属性_> numberFormatLocal|用户语言字符串形式表示给定范围的 Excel 的数字格式代码。|1.7|
|[range](/javascript/api/excel/excel.range)|_属性_ > style|代表当前范围的样式。 这将返回 null 或 string。|1.7|
|[range](/javascript/api/excel/excel.range)|_方法_> getAbsoluteResizedRange (numRows： 数字、 numColumns： 数)|获取一个 Range 对象，具有相同的左上角单元格作为当前的 Range 对象，但是的行和列的指定数字。|1.7|
|[range](/javascript/api/excel/excel.range)|_方法_> getImage()|为 base64 编码的图像呈现范围。|1.7|
|[range](/javascript/api/excel/excel.range)|_方法_> getSurroundingRegion()|返回一个 Range 对象，该对象表示的此区域的左上角单元格周围的区域。 周围的区域是受空行和空列相对于此范围的任意组合的范围。|1.7|
|[range](/javascript/api/excel/excel.range)|_方法_> showCard()|如果它具有丰富的价值内容，显示为活动单元格卡片。|1.7|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_属性_> textOrientation|获取或设置范围内的所有单元格的文本方向。|1.7|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_属性_> useStandardHeight|确定 Range 对象的行高是否等于工作表的标准高度。|1.7|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_属性_> useStandardWidth|确定 Range 对象的 columnwidth 是否等于的工作表的标准列宽。|1.7|
|[rangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|_属性_ > address|代表超链接的 url 目标。|1.7|
|[rangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|_属性_> 文档.|表示的文档。 超链接的目标。|1.7|
|[rangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|_属性_> 屏幕提示|表示悬停超链接时显示的字符串。|1.7|
|[rangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|_属性_> textToDisplay|代表左上区域中的大多数单元格中显示的字符串。|1.7|
|[style](/javascript/api/excel/excel.style)|_属性_> addIndent|指示是否文本会自动缩进单元格中的文本对齐方式设置为等距分布时。|1.7|
|[style](/javascript/api/excel/excel.style)|_属性_> 自动|指示是否文本会自动缩进单元格中的文本对齐方式设置为等距分布时。|1.7|
|[style](/javascript/api/excel/excel.style)|_属性_> builtIn|指示是否样式为内置样式。 只读。|1.7|
|[style](/javascript/api/excel/excel.style)|_属性_ > formulaHidden|指示是否将在工作表处于保护状态时隐藏公式。|1.7|
|[style](/javascript/api/excel/excel.style)|_属性_> horizontalAlignment|代表样式的水平对齐方式。 可能的值为： 常规，左、 居中、 右、 填充、 两端对齐、 CenterAcrossSelection、 分布式。|1.7|
|[style](/javascript/api/excel/excel.style)|_属性_> includeAlignment|指示该样式是否包括自动、 HorizontalAlignment、 VerticalAlignment、 WrapText、 IndentLevel 和 TextOrientation 属性。|1.7|
|[style](/javascript/api/excel/excel.style)|_属性_> includeBorder|指示该样式是否包括的颜色、 ColorIndex、 LineStyle 和权重边框属性。|1.7|
|[style](/javascript/api/excel/excel.style)|_属性_> includeFont|指示该样式是否包括的背景、 加粗、 颜色、 ColorIndex、 FontStyle、 斜体、 名称、 大小、 删除线、 下标、 上标和下划线的字体属性。|1.7|
|[style](/javascript/api/excel/excel.style)|_属性_> includeNumber|指示该样式是否包括 NumberFormat 属性。|1.7|
|[style](/javascript/api/excel/excel.style)|_属性_> includePatterns|指示该样式是否包括颜色、 ColorIndex、 InvertIfNegative、 模式、 PatternColor 和 PatternColorIndex interior 属性。|1.7|
|[style](/javascript/api/excel/excel.style)|_属性_> includeProtection|指示该样式是否包括 FormulaHidden 和锁定保护属性。|1.7|
|[style](/javascript/api/excel/excel.style)|_属性_> indentLevel|一个介于 0 至 250 个的整数，指示样式的缩进级别。|1.7|
|[style](/javascript/api/excel/excel.style)|_属性_ > locked|指示是否对象已锁定时的工作表处于保护。|1.7|
|[style](/javascript/api/excel/excel.style)|_属性_ > name|样式的名称。 只读。|1.7|
|[style](/javascript/api/excel/excel.style)|_属性_ > numberFormat|样式的数字格式的格式代码。|1.7|
|[style](/javascript/api/excel/excel.style)|_属性_> numberFormatLocal|样式的数字格式本地化的格式代码。|1.7|
|[style](/javascript/api/excel/excel.style)|_属性_> 方向|样式的文本方向。|1.7|
|[style](/javascript/api/excel/excel.style)|_属性_> readingOrder|样式的读取顺序。 可能的值为： 上下文，LeftToRight，RightToLeft。|1.7|
|[style](/javascript/api/excel/excel.style)|_属性_> shrinkToFit|指示文本是否自动缩小以适合可用列宽。|1.7|
|[style](/javascript/api/excel/excel.style)|_属性_> textOrientation|样式的文本方向。|1.7|
|[style](/javascript/api/excel/excel.style)|_属性_> verticalAlignment|代表样式的垂直对齐方式。 可能的值为： 顶部中心、 底部、 两端对齐、 分布式。|1.7|
|[style](/javascript/api/excel/excel.style)|_属性_> wrapText|指示是否 Microsoft Excel 会将文本包装对象中。|1.7|
|[style](/javascript/api/excel/excel.style)|_关系_> 边框|边框表示的四个边框样式的四个 Border 对象的集合。 只读。|1.7|
|[style](/javascript/api/excel/excel.style)|_关系_> 填充|样式的填充。 只读。|1.7|
|[style](/javascript/api/excel/excel.style)|_关系_ > font|Font 对象，代表样式的字体。 只读。|1.7|
|[style](/javascript/api/excel/excel.style)|_方法_> delete()|删除此样式。|1.7|
|[styleCollection](/javascript/api/excel/excel.stylecollection)|_属性_ > items|Style 对象的集合。 只读。|1.7|
|[styleCollection](/javascript/api/excel/excel.stylecollection)|_方法_> add(name: string)]|向集合添加新的样式。|1.7|
|[styleCollection](/javascript/api/excel/excel.stylecollection)|_方法_> getItem(name: string)|按名称获取样式。|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_属性_ > address|获取表示特定工作表的表的更改的区域的地址。|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_属性_> changeType|获取表示 Changed 事件触发的方式的更改类型。 可能的值为： 其他人，RangeEdited、 RowInserted、 RowDeleted、 ColumnInserted、 ColumnDeleted、 CellInserted、 CellDeleted。|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_属性_> 源|获取在事件的源。 可能的值为： 本地、 远程。|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_属性_> tableId|获取更改的数据的表的 id。|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_属性_> 类型|获取该事件的类型。 可能的值为： WorksheetDataChanged、 WorksheetSelectionChanged、 WorksheetAdded、 WorksheetActivated、 WorksheetDeactivated、 TableDataChanged、 TableSelectionChanged、 WorksheetDeleted。|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_属性_> worksheetId|获取更改的数据在工作表的 id。|1.7|
|[tableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|_属性_ > address|获取代表特定工作表上表的选定的区域的范围地址。|1.7|
|[tableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|_属性_> isInsideTable|指示是否所选内容位于表格，地址会导致无用如果 IsInsideTable 为 false。|1.7|
|[tableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|_属性_> tableId|获取的表中所选内容更改的 id。|1.7|
|[tableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|_属性_> 类型|获取该事件的类型。 可能的值为： WorksheetDataChanged、 WorksheetSelectionChanged、 WorksheetAdded、 WorksheetActivated、 WorksheetDeactivated、 TableDataChanged、 TableSelectionChanged、 WorksheetDeleted。|1.7|
|[tableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|_属性_> worksheetId|获取工作表中所选内容更改的 id。|1.7|
|[workbook](/javascript/api/excel/excel.workbook)|_属性_ > name|获取工作簿的名称。 只读。|1.7|
|[workbook](/javascript/api/excel/excel.workbook)|_关系_> dataConnections|刷新工作簿中的所有数据连接。 只读。|1.7|
|[workbook](/javascript/api/excel/excel.workbook)|_关系_ > properties|获取工作簿属性。 只读。|1.7|
|[workbook](/javascript/api/excel/excel.workbook)|_关系_ > protection|返回工作簿的 workbook protection 对象。 只读。|1.7|
|[workbook](/javascript/api/excel/excel.workbook)|_关系_> 样式|代表工作簿相关联的样式的集合。 只读。|1.7|
|[workbook](/javascript/api/excel/excel.workbook)|_方法_> getActiveCell()|从工作簿中获取当前活动单元格。|1.7|
|[workbookProtection](/javascript/api/excel/excel.workbookprotection)|_属性_ > protected|指示是否保护工作簿。 只读。 只读。|1.7|
|[workbookProtection](/javascript/api/excel/excel.workbookprotection)|_方法_> protect(password: string)|保护工作簿。 如果工作簿处于保护状态，失败。|1.7|
|[workbookProtection](/javascript/api/excel/excel.workbookprotection)|_方法_> unprotect(password: string)|取消工作簿的保护。|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_属性_> 网格线|获取或设置工作表的网格线标志。|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_属性_> 标题|获取或设置工作表的标题的标志。|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_属性_> showHeadings|获取或设置工作表的标题的标志。|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_属性_> standardHeight|在表中，以磅为单位返回 standard （默认值） 所有行的高度。 只读。|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_属性_> standardWidth|返回或设置工作表中的所有列的标准 （默认值） 宽度。|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_属性_> tabColor|获取或设置工作表标签颜色。|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_关系_> freezePanes|获取一个对象，可用来操作工作表上的冻结的窗格只读的。|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_方法_> 副本 (positionType: WorksheetPositionType，relativeTo： 工作表)|复制工作表，并将其放置在指定的位置。 返回复制的工作表。|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_方法_> getRangeByIndexes (startRow： 号码后，startColumn： 号码后，rowCount： 号码后，columnCount： 数)| 获取以特定行索引和列索引开始并跨越了一定数量的行和列的 range 对象。|1.7|
|[worksheetActivatedEventArgs](/javascript/api/excel/excel.worksheetactivatedeventargs)|_属性_> 类型|获取该事件的类型。 可能的值为： WorksheetDataChanged、 WorksheetSelectionChanged、 WorksheetAdded、 WorksheetActivated、 WorksheetDeactivated、 TableDataChanged、 TableSelectionChanged、 WorksheetDeleted。|1.7|
|[worksheetActivatedEventArgs](/javascript/api/excel/excel.worksheetactivatedeventargs)|_属性_> worksheetId|获取已激活的工作表的 id。|1.7|
|[worksheetAddedEventArgs](/javascript/api/excel/excel.worksheetaddedeventargs)|_属性_> 源|获取在事件的源。 可能的值为： 本地、 远程。|1.7|
|[worksheetAddedEventArgs](/javascript/api/excel/excel.worksheetaddedeventargs)|_属性_> 类型|获取该事件的类型。 可能的值为： WorksheetDataChanged、 WorksheetSelectionChanged、 WorksheetAdded、 WorksheetActivated、 WorksheetDeactivated、 TableDataChanged、 TableSelectionChanged、 WorksheetDeleted。|1.7|
|[worksheetAddedEventArgs](/javascript/api/excel/excel.worksheetaddedeventargs)|_属性_> worksheetId|获取添加到工作簿的工作表的 id。|1.7|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|_属性_ > address|获取代表特定工作表的更改区域的范围地址。|1.7|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|_属性_> changeType|获取表示 Changed 事件触发的方式的更改类型。 可能的值为： 其他人，RangeEdited、 RowInserted、 RowDeleted、 ColumnInserted、 ColumnDeleted、 CellInserted、 CellDeleted。|1.7|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|_属性_> 源|获取在事件的源。 可能的值为： 本地、 远程。|1.7|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|_属性_> 类型|获取该事件的类型。 可能的值为： WorksheetDataChanged、 WorksheetSelectionChanged、 WorksheetAdded、 WorksheetActivated、 WorksheetDeactivated、 TableDataChanged、 TableSelectionChanged、 WorksheetDeleted。|1.7|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|_属性_> worksheetId|获取更改的数据在工作表的 id。|1.7|
|[worksheetDeactivatedEventArgs](/javascript/api/excel/excel.worksheetdeactivatedeventargs)|_属性_> 类型|获取该事件的类型。 可能的值为： WorksheetDataChanged、 WorksheetSelectionChanged、 WorksheetAdded、 WorksheetActivated、 WorksheetDeactivated、 TableDataChanged、 TableSelectionChanged、 WorksheetDeleted。|1.7|
|[worksheetDeactivatedEventArgs](/javascript/api/excel/excel.worksheetdeactivatedeventargs)|_属性_> worksheetId|获取的停用工作表的 id。|1.7|
|[worksheetDeletedEventArgs](/javascript/api/excel/excel.worksheetdeletedeventargs)|_属性_> 源|获取在事件的源。 可能的值为： 本地、 远程。|1.7|
|[worksheetDeletedEventArgs](/javascript/api/excel/excel.worksheetdeletedeventargs)|_属性_> 类型|获取该事件的类型。 可能的值为： WorksheetDataChanged、 WorksheetSelectionChanged、 WorksheetAdded、 WorksheetActivated、 WorksheetDeactivated、 TableDataChanged、 TableSelectionChanged、 WorksheetDeleted。|1.7|
|[worksheetDeletedEventArgs](/javascript/api/excel/excel.worksheetdeletedeventargs)|_属性_> worksheetId|获取从工作簿中删除的工作表的 id。|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_方法_> freezeAt (frozenRange： 范围或字符串)|活动工作表视图中设置冻结单元格。|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_方法_> freezeColumns(count: number)|冻结就地工作表的第一列。|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_方法_> freezeRows(count: number)|冻结就地工作表的顶部行。|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_方法_> getLocation()|获取介绍活动工作表视图中的冻结单元格的范围。|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_方法_> getLocationOrNullObject()|获取介绍活动工作表视图中的冻结单元格的范围。|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_方法_> unfreeze()|工作表中删除所有的冻结的窗格。|1.7|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_属性_> allowEditObjects|代表工作表保护选项的允许编辑对象。|1.7|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_属性_> allowEditScenarios|表示允许编辑方案的工作表保护选项。|1.7|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_关系_> selectionMode|表示选定模式的工作表保护选项。|1.7|
|[worksheetSelectionChangedEventArgs](/javascript/api/excel/excel.worksheetselectionchangedeventargs)|_属性_ > address|获取代表特定工作表的选定的区域的范围地址。|1.7|
|[worksheetSelectionChangedEventArgs](/javascript/api/excel/excel.worksheetselectionchangedeventargs)|_属性_> 类型|获取该事件的类型。 可能的值为： WorksheetDataChanged、 WorksheetSelectionChanged、 WorksheetAdded、 WorksheetActivated、 WorksheetDeactivated、 TableDataChanged、 TableSelectionChanged、 WorksheetDeleted。|1.7|
|[worksheetSelectionChangedEventArgs](/javascript/api/excel/excel.worksheetselectionchangedeventargs)|_属性_> worksheetId|获取工作表中所选内容更改的 id。|1.7|


## <a name="whats-new-in-excel-javascript-api-16"></a>What's new in Excel JavaScript API 1.6 

### <a name="conditional-formatting"></a>条件格式

介绍区域的条件格式。 允许以下类型的条件格式：

* 彩色温标
* 数据栏
* 图标集
* 自定义

此外：

* 返回条件格式应用于的区域。 
* 条件格式的删除。 
* 提供了优先级和 stopifTrue 功能。 
* 获取给定范围内所有条件格式的集合。 
* 清除当前指定范围中处于活动状态的所有条件格式。 

|对象| 最近更新| 说明|要求集|
|:----|:----|:----|:----|
|[application](/javascript/api/excel/excel.application)|_方法_> suspendApiCalculationUntilNextSync()|在下一次调用“context.sync()”前暂停计算。设置后，开发者负责重新计算工作簿，以确保传播所有依赖项。|1.6|
|[cellValueConditionalFormat](/javascript/api/excel/excel.cellvalueconditionalformat)|_关系_> 格式|返回一个格式对象，其中封装了条件格式字体、填充、边框和其他属性。只读。|1.6|
|[cellValueConditionalFormat](/javascript/api/excel/excel.cellvalueconditionalformat)|_关系_> 规则|表示此条件格式中的 Rule 对象。|1.6|
|[colorScaleConditionalFormat](/javascript/api/excel/excel.colorscaleconditionalformat)|_属性_> threeColorScale|如果为 true，则彩色温标有三个点（最小、中点、最大），否则将有两个点（最小、最大）。只读。|1.6|
|[colorScaleConditionalFormat](/javascript/api/excel/excel.colorscaleconditionalformat)|_关系_ > criteria|彩色温标的条件。使用两点彩色温标时，中点可选。|1.6|
|[conditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule)|_属性_> formula1|如果需要，公式可对条件格式规则进行求值。|1.6|
|[conditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule)|_属性_> formula2|如果需要，公式可对条件格式规则进行求值。|1.6|
|[conditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule)|_属性_ > operator|文本格式条件的运算符。可能的值是：Invalid、Between、NotBetween、EqualTo、NotEqualTo、GreaterThan、LessThan、GreaterThanOrEqual、LessThanOrEqual。|1.6|
|[conditionalColorScaleCriteria](/javascript/api/excel/excel.conditionalcolorscalecriteria)|_关系_> 最大|最大点彩色温标条件。|1.6|
|[conditionalColorScaleCriteria](/javascript/api/excel/excel.conditionalcolorscalecriteria)|_关系_> 中点|彩色温标为 3 色温标时的中点彩色温标条件。|1.6|
|[conditionalColorScaleCriteria](/javascript/api/excel/excel.conditionalcolorscalecriteria)|_关系_> 最低|最小点彩色温标条件。|1.6|
|[conditionalColorScaleCriterion](/javascript/api/excel/excel.conditionalcolorscalecriterion)|_属性_ > color|彩色温标颜色的 HTML 颜色代码表示。例如，#FF0000 表示红色。|1.6|
|[conditionalColorScaleCriterion](/javascript/api/excel/excel.conditionalcolorscalecriterion)|_属性_> 公式|数字、公式或 null（如果类型为 LowestValue）。|1.6|
|[conditionalColorScaleCriterion](/javascript/api/excel/excel.conditionalcolorscalecriterion)|_属性_> 类型|应基于的图标条件公式。可能的值是：Invalid、LowestValue、HighestValue、Number、Percent、Formula、Percentile。|1.6|
|[conditionalDataBarNegativeFormat](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|_属性_> borderColor|表示窗体 #RRGGBB（例如“FFA500”）的边框线条颜色或作为已命名的 HTML 颜色（例如“orange”）的 HTML 颜色代码。|1.6|
|[conditionalDataBarNegativeFormat](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|_属性_> fillColor|表示窗体 #RRGGBB（例如 "FFA500"）的填充颜色或作为已命名的 HTML 颜色（例如 "orange"）的 HTML 颜色代码。|1.6|
|[conditionalDataBarNegativeFormat](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|_属性_> matchPositiveBorderColor|负 DataBar 是否与正 DataBar 具有相同边框颜色的布尔值表示形式。|1.6|
|[conditionalDataBarNegativeFormat](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|_属性_> matchPositiveFillColor|负 DataBar 是否与正 DataBar 具有相同填充颜色的布尔值表示形式。|1.6|
|[conditionalDataBarPositiveFormat](/javascript/api/excel/excel.conditionaldatabarpositiveformat)|_属性_> borderColor|表示窗体 #RRGGBB（例如“FFA500”）的边框线条颜色或作为已命名的 HTML 颜色（例如“orange”）的 HTML 颜色代码。|1.6|
|[conditionalDataBarPositiveFormat](/javascript/api/excel/excel.conditionaldatabarpositiveformat)|_属性_> fillColor|表示窗体 #RRGGBB（例如 "FFA500"）的填充颜色或作为已命名的 HTML 颜色（例如 "orange"）的 HTML 颜色代码。|1.6|
|[conditionalDataBarPositiveFormat](/javascript/api/excel/excel.conditionaldatabarpositiveformat)|_属性_> gradientFill|DataBar 是否具有渐变的布尔值表示形式。|1.6|
|[conditionalDataBarRule](/javascript/api/excel/excel.conditionaldatabarrule)|_属性_> 公式|如果需要，公式可对 databar 规则进行求值。|1.6|
|[conditionalDataBarRule](/javascript/api/excel/excel.conditionaldatabarrule)|_属性_> 类型|Databar 的规则的类型。可能的值是：LowestValue、HighestValue、Number、Percent、Formula、Percentile、Automatic。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_属性_ > id|条件格式中当前 ConditionalFormatCollection 优先级。 只读。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_属性_> 优先级|优先级（或索引）位于此条件格式当前存在的条件格式集合内。也会更改此处|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_属性_> stopIfTrue|如果满足此条件格式的条件，则不会有任何低优先级格式应在此单元格上生效。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_属性_> 类型|条件格式的类型。一次仅可设置一个。只读。只读。可能的值是：“Custom”、“DataBar”、“ColorScale”、“IconSet”。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_关系_> cellValue|如果当前的条件格式是 CellValue 类型，则返回单元值条件格式属性。只读。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_关系_> cellValueOrNullObject|如果当前的条件格式是 CellValue 类型，则返回单元值条件格式属性。只读。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_关系_> colorScale|如果当前的条件格式为 ColorScale 类型，返回 ColorScale 条件格式属性。只读。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_关系_> colorScaleOrNullObject|如果当前的条件格式为 ColorScale 类型，返回 ColorScale 条件格式属性。只读。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_关系_> 自定义|如果当前的条件格式是自定义类型，则返回自定义条件格式属性。只读。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_关系_> customOrNullObject|如果当前的条件格式是自定义类型，则返回自定义条件格式属性。只读。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_关系_> dataBar|如果当前的条件格式是数据栏，则返回数据栏属性。只读。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_关系_> dataBarOrNullObject|如果当前的条件格式是数据栏，则返回数据栏属性。只读。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_关系_> iconSet|如果当前的条件格式是 IconSet 类型，则返回 IconSet 条件格式属性。只读。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_关系_> iconSetOrNullObject|如果当前的条件格式是 IconSet 类型，则返回 IconSet 条件格式属性。只读。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_关系_> 预设|返回预设条件的条件格式，如上述 averagebelow averageunique valuescontains blanknonblankerrornoerror 属性。只读。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_关系_> presetOrNullObject|返回预设条件的条件格式，如上述 averagebelow averageunique valuescontains blanknonblankerrornoerror 属性。只读。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_关系_> textComparison|如果当前的条件格式是文本类型，则返回特定文本条件格式属性。只读。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_关系_> textComparisonOrNullObject|如果当前的条件格式是文本类型，则返回特定文本条件格式属性。只读。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_关系_> 从上到下|如果当前的条件格式是 TopBottom 类型，则返回 TopBottom 条件格式属性。只读。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_关系_> topBottomOrNullObject|如果当前的条件格式是 TopBottom 类型，则返回 TopBottom 条件格式属性。只读。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_方法_> delete()|删除此条件格式。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_方法_> getRange()|返回条件格式应用的区域，如果区域不连续，则返回 NULL 对象。只读。|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_方法_> getRangeOrNullObject()|返回条件格式应用的区域，如果区域不连续，则返回 NULL 对象。只读。|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_属性_ > items|ConditionalFormat 对象的集合。只读。|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_方法_> add(type: string)|向 firsttop 优先级的集合添加新的条件格式。|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_方法_> clearAll()|清除当前指定范围中处于活动状态的所有条件格式。|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_方法_> getCount()|返回工作簿中的条件格式数。只读。|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_方法_> getItem(id: string)|条件格式返回给定的 id。|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_方法_> getItemAt(index: number)|返回给定索引处的条件格式。|1.6|
|[conditionalFormatRule](/javascript/api/excel/excel.conditionalformatrule)|_属性_> 公式|如果需要，公式可对条件格式规则进行求值。|1.6|
|[conditionalFormatRule](/javascript/api/excel/excel.conditionalformatrule)|_属性_> formulaLocal|如果需要，公式可采用用户的语言对条件格式规则进行求值。|1.6|
|[conditionalFormatRule](/javascript/api/excel/excel.conditionalformatrule)|_属性_> formulaR1C1|如果需要，公式可采用 R1C1 表示法对条件格式规则进行求值。|1.6|
|[conditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|_属性_> 公式|取决于类型的数字或公式。|1.6|
|[conditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|_属性_ > operator|Icon 条件格式的每个规则类型的 GreaterThan 或 GreaterThanOrEqual。可能的值是：Invalid、GreaterThan、GreaterThanOrEqual。|1.6|
|[conditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|_关系_> customIcon|如果与默认 IconSet 不同，返回当前条件的自定义图标，否则将返回 null。|1.6|
|[conditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|_关系_ > type|应基于的图标条件公式。|1.6|
|[conditionalPresetCriteriaRule](/javascript/api/excel/excel.conditionalpresetcriteriarule)|_属性_> 条件|条件格式的条件。可能的值是：Invalid、Blanks、NonBlanks、Errors、NonErrors、Yesterday、Today、Tomorrow、LastSevenDays、LastWeek、ThisWeek、NextWeek、LastMonth、ThisMonth、NextMonth、AboveAverage、BelowAverage、EqualOrAboveAverage、EqualOrBelowAverage、OneStdDevAboveAverage、OneStdDevBelowAverage、TwoStdDevAboveAverage、TwoStdDevBelowAverage、ThreeStdDevAboveAverage、ThreeStdDevBelowAverage、UniqueValues、DuplicateValues。|1.6|
|[conditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|_属性_ > color|表示窗体 #RRGGBB（例如“FFA500”）的边框线条颜色或作为已命名的 HTML 颜色（例如“orange”）的 HTML 颜色代码。|1.6|
|[conditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|_属性_ > id|表示边框标识符。只读。可能的值是：EdgeTop、EdgeBottom、EdgeLeft、EdgeRight。|1.6|
|[conditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|_属性_> sideIndex|指示边框的特定边的常量值。只读。可能的值是：EdgeTop、EdgeBottom、EdgeLeft、EdgeRight。|1.6|
|[conditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|_属性_ > style|线条样式的常量之一，指定边框的线条样式。可能的值是：None、Continuous、Dash、DashDot、DashDotDot、Dot、Double、SlantDashDot。|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_属性_> 计数|集合中的 border 对象数量。只读。|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_属性_ > items|conditionalRangeBorder 对象的集合。只读。|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_关系_> 底部|获取只读的顶部边框。|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_关系_> 左侧|获取只读的顶部边框。|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_关系_> 右|获取只读的顶部边框。|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_关系_> 顶部|获取只读的顶部边框。|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_方法_> getItem(index: string)|使用其名称获取 border 对象|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_方法_> getItemAt(index: number)|使用其索引获取 border 对象|1.6|
|[conditionalRangeFill](/javascript/api/excel/excel.conditionalrangefill)|_属性_ > color|表示窗体 #RRGGBB（例如 "FFA500"）的填充颜色或作为已命名的 HTML 颜色（例如 "orange"）的 HTML 颜色代码。|1.6|
|[conditionalRangeFill](/javascript/api/excel/excel.conditionalrangefill)|_方法_> clear()|重置填充。|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_属性_> 加粗|表示字体的加粗状态。|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_属性_ > color|文本颜色的 HTML 颜色代码表示。例如，#FF0000 表示红色。|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_属性_> 斜体|表示字体的斜体状态。|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_属性_> 删除线|表示字体的删除线状态。|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_属性_> 下划线|应用于字体的下划线类型。可能的值是：None、Single、Double。|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_方法_> clear()|重置字体格式。|1.6|
|[conditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|_属性_ > numberFormat|表示 Excel 中指定范围的数字格式代码。当传递 null 时清除。|1.6|
|[conditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|_关系_> 边框|应用于整个条件格式范围的 border 对象的集合。只读。|1.6|
|[conditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|_关系_> 填充|返回在整个条件格式范围内定义的 fill 对象。只读。|1.6|
|[conditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|_关系_ > font|返回在整个条件格式范围内定义的 font 对象。只读。|1.6|
|[conditionalTextComparisonRule](/javascript/api/excel/excel.conditionaltextcomparisonrule)|_属性_ > operator|文本格式条件的运算符。可能的值是：Invalid、Contains、NotContains、BeginsWith、EndsWith。|1.6|
|[conditionalTextComparisonRule](/javascript/api/excel/excel.conditionaltextcomparisonrule)|_属性_ > text|条件格式的文本值。|1.6|
|[conditionalTopBottomRule](/javascript/api/excel/excel.conditionaltopbottomrule)|_属性_> 排名|1 和 1000 之间的数字排名或 1 和 100 之间的百分比排名。|1.6|
|[conditionalTopBottomRule](/javascript/api/excel/excel.conditionaltopbottomrule)|_属性_> 类型|基于排名第一或排名最后的格式值。可能的值是：Invalid、TopItems、TopPercent、BottomItems、BottomPercent。|1.6|
|[customConditionalFormat](/javascript/api/excel/excel.customconditionalformat)|_关系_> 格式|返回一个格式对象，其中封装了条件格式字体、填充、边框和其他属性。只读。|1.6|
|[customConditionalFormat](/javascript/api/excel/excel.customconditionalformat)|_关系_> 规则|表示此条件格式中的 Rule 对象。只读。|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_属性_> axisColor|表示窗体 #RRGGBB（例如 "FFA500"）的轴行颜色或作为已命名的 HTML 颜色（例如 "orange"）的 HTML 颜色代码。|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_属性_> axisFormat|如何确定 Excel 数据栏的轴的表示形式。可能的值是：Automatic、None、CellMidPoint。|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_属性_> barDirection|表示数据栏图形应遵循的方向。可能的值是：Context、LeftToRight、RightToLeft。|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_属性_> showDataBarOnly|如果为 true，则对应用数据栏的单元格隐藏值。|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_关系_> lowerBoundRule|构成数据栏的下限（以及如何计算，如果适用）的规则。|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_关系_> negativeFormat|Excel 数据栏中轴左侧的所有值的表示形式。只读。|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_关系_> positiveFormat|Excel 数据栏中轴右侧的所有值的表示形式。只读。|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_关系_> upperBoundRule|构成数据栏的上限（以及如何计算，如果适用）的规则。|1.6|
|[iconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|_属性_> reverseIconOrder|如果为 true，则反转 IconSet 的图标顺序。注意，如果使用自定义图标，则不能进行设置。|1.6|
|[iconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|_属性_> showIconOnly|如果为 true，则隐藏值并仅显示图标。|1.6|
|[iconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|_属性_ > style|如果设置，则显示条件格式的 IconSet 选项。可能的值是：Invalid、ThreeArrows、ThreeArrowsGray、ThreeFlags、ThreeTrafficLights1、ThreeTrafficLights2、ThreeSigns、ThreeSymbols、ThreeSymbols2、FourArrows、FourArrowsGray、FourRedToBlack、FourRating、FourTrafficLights、FiveArrows、FiveArrowsGray、FiveRating、FiveQuarters、ThreeStars、ThreeTriangles、FiveBoxes。|1.6|
|[iconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|_关系_ > criteria|规则的 Criteria 和 IconSet 数组，以及条件图标的潜在自定义图标。注意，对于第一个条件，只能修改自定义图标，类型、公式和运算符在设置时将忽略。|1.6|
|[presetCriteriaConditionalFormat](/javascript/api/excel/excel.presetcriteriaconditionalformat)|_关系_> 格式|返回一个格式对象，其中封装了条件格式字体、填充、边框和其他属性。只读。|1.6|
|[presetCriteriaConditionalFormat](/javascript/api/excel/excel.presetcriteriaconditionalformat)|_关系_> 规则|条件格式的规则。|1.6|
|[range](/javascript/api/excel/excel.range)|_关系_> conditionalFormats|范围交叉的 ConditionalFormats 的集合。只读。|1.6|
|[range](/javascript/api/excel/excel.range)|_方法_> calculate()|计算工作表上的单元格区域。|1.6|
|[textConditionalFormat](/javascript/api/excel/excel.textconditionalformat)|_关系_> 格式|返回一个格式对象，其中封装了条件格式字体、填充、边框和其他属性。只读。|1.6|
|[textConditionalFormat](/javascript/api/excel/excel.textconditionalformat)|_关系_> 规则|条件格式的规则。|1.6|
|[topBottomConditionalFormat](/javascript/api/excel/excel.topbottomconditionalformat)|_关系_> 格式|返回一个格式对象，其中封装了条件格式字体、填充、边框和其他属性。只读。|1.6|
|[topBottomConditionalFormat](/javascript/api/excel/excel.topbottomconditionalformat)|_关系_> 规则|表示 TopBottom 条件格式的条件。|1.6|
|[workbook](/javascript/api/excel/excel.workbook)|_关系_> internalTest|仅供内部使用。只读。|1.6|
|[worksheet](/javascript/api/excel/excel.worksheet)|_方法_> calculate(markAllDirty: bool)|计算工作表上的所有单元格。|1.6|

##  <a name="whats-new-in-excel-javascript-api-15"></a>What's new in Excel JavaScript API 1.5

### <a name="custom-xml-part"></a>自定义 XML 部件

* 将自定义 XML 部件集合添加到工作簿对象中。
* 使用 ID 获取自定义 XML 部件
* 获取其命名空间匹配给定命名空间的自定义 XML 部件的新作用域内集合。
* 获取与某个部件关联的 XML 字符串。
* 提供部件的 ID 和命名空间。
* 向工作簿添加新的自定义 XML 部件。
* 设置整个 XML 部件。
* 删除自定义 XML 部件。
* 删除其给定名称来自由 xpath 标识的元素的属性。
* 按 xpath 查询 XML 内容。
* 插入、更新和删除属性。

**参考实现：** 请参阅[此处](https://github.com/mandren/Excel-CustomXMLPart-Demo)，了解说明如何在外接程序中使用自定义 XML 部件的参考实现。

### <a name="others"></a>其他
* `range.getSurroundingRegion()` 返回一个 Range 对象，该对象表示此范围的周围区域。周围区域是由相对于该范围的空白行和空白列的任何组合所限定的范围。
* 对表列执行 `getNextColumn()`、`getPreviousColumn()` 以及 `getLast() 操作。
* 对工作簿执行 `getActiveWorksheet()` 操作。
* 工作簿的 `getRange(address: string)` 关闭。
* `getBoundingRange(ranges: )` 获取包含提供的范围的最小 range 对象。例如，介于 “B2:C5” 和 “D10:E15” 之间的边界范围为 “B2:E15”。
* 对各种集合（例如已命名项目、工作表、表等）执行 `getCount()` 操作以获取集合中的项目数。 `workbook.worksheets.getCount()`
* 对各种集合（如工作表、表列、图标点、范围视图集合）执行 `getFirst()` 和 `getLast()` 以及 get last 操作。
* 对工作表、表列集合执行 `getNext()` 和 `getPrevious()` 操作。
* `getRangeR1C1()` 获取以特定行索引和列索引开始并跨越了一定数量的行和列的 range 对象。

|对象| 最近更新| 说明|要求集|
|:----|:----|:----|:----|
|[customXmlPart](/javascript/api/excel/excel.customxmlpart)|_属性_ > id|自定义 XML 部件的 ID。只读。|1.5|
|[customXmlPart](/javascript/api/excel/excel.customxmlpart)|_属性_> namespaceUri|自定义 XML 部件的命名空间 URI。只读。|1.5|
|[customXmlPart](/javascript/api/excel/excel.customxmlpart)|_方法_> delete()|删除自定义 XML 部件。|1.5|
|[customXmlPart](/javascript/api/excel/excel.customxmlpart)|_方法_> getXml()|获取自定义 XML 部件的完整 XML 内容。|1.5|
|[customXmlPart](/javascript/api/excel/excel.customxmlpart)|_方法_> setXml(xml: string)|设置自定义 XML 部件的完整 XML 内容。|1.5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_属性_ > items|customXmlPart 对象的集合。只读。|1.5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_方法_> add(xml: string)|向工作簿添加新的自定义 XML 部件。|1.5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_方法_> getByNamespace(namespaceUri: string)|获取其命名空间匹配给定命名空间的自定义 XML 部件的新作用域内集合。|1.5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_方法_> getCount()|获取此集合中 CustomXml 部件的数量。|1.5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_方法_> getItem(id: string)|获取基于其 ID 的自定义 XML 部件。|1.5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_方法_> getItemOrNullObject(id: string)|获取基于其 ID 的自定义 XML 部件。|1.5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_属性_ > items|CustomXmlPartScoped 对象的集合。只读。|1.5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_方法_> getCount()|获取此集合中 CustomXML 部件的数量。|1.5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_方法_> getItem(id: string)|获取基于其 ID 的自定义 XML 部件。|1.5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_方法_> getItemOrNullObject(id: string)|获取基于其 ID 的自定义 XML 部件。|1.5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_方法_> getOnlyItem()|如果集合仅包含一个项，则此方法返回该项。|1.5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_方法_> getOnlyItemOrNullObject()|如果集合仅包含一个项，则此方法返回该项。|1.5|
|[workbook](/javascript/api/excel/excel.workbook)|_关系_> customXmlParts|代表此工作簿包含的自定义 XML 部件的集合。只读。|1.5|
|[worksheet](/javascript/api/excel/excel.worksheet)|_方法_> getNext(visibleOnly: bool)|获取该工作表之后的工作表。如果该工作表后没有工作表，此方法将引发错误。|1.5|
|[worksheet](/javascript/api/excel/excel.worksheet)|_方法_> getNextOrNullObject(visibleOnly: bool)|获取该工作表之后的工作表。如果该工作表后没有工作表，此方法将返回 null 对象。|1.5|
|[worksheet](/javascript/api/excel/excel.worksheet)|_方法_> getPrevious(visibleOnly: bool)|获取该工作表之前的工作表。如果没有以前的工作表，此方法将引发错误。|1.5|
|[worksheet](/javascript/api/excel/excel.worksheet)|_方法_> getPreviousOrNullObject(visibleOnly: bool)|获取该工作表之前的工作表。如果没有以前的工作表，此方法将返回 null 对象。|1.5|
|[worksheetCollection](/javascript/api/excel/excel.worksheetcollection)|_方法_> getFirst(visibleOnly: bool)|获取集合中的第一个工作表。|1.5|
|[worksheetCollection](/javascript/api/excel/excel.worksheetcollection)|_方法_> getLast(visibleOnly: bool)|获取集合中的最后一个工作表。|1.5|

## <a name="whats-new-in-excel-javascript-api-14"></a>Excel JavaScript API 1.4 的最近更新
以下是新增功能要求中的 Excel JavaScript Api 设置 1.4。

### <a name="named-item-add-and-new-properties"></a>添加了已命名项和新属性

新属性：

* `comment`
* `scope`：限定到工作表或工作簿的项。
* `worksheet`：返回已命名项限定到的工作表。

新方法：

* `add(name: string, reference: Range or string, comment: string)`：将新名称添加到给定范围的集合。
* `addFormulaLocal(name: string, formula: string, comment: string)`：使用用户的公式区域设置，将新名称添加到给定范围的集合。

### <a name="settings-api-in-in-excel-namespace"></a>Excel 命名空间中的设置 API

[Setting](/javascript/api/excel/excel.setting) 对象表示文档保留设置的键值对。现在，我们已在 Excel 命名空间下添加了与设置相关的 API。这不会提供全新功能，但可便于继续使用基于承诺的批处理 API 语法，减少对 Excel 相关任务常见 API 的依赖。

API 包括通过键获取设置条目的 `getItem()`，以及将指定键值设置对添加到工作簿的 `add()`。

### <a name="others"></a>其他

* 设置表列名称（旧版只允许读取）。
* 将表列添加到表的末尾（旧版只允许添加到除末尾之外的其他任何位置）。
* 一次性向表中添加多行（旧版只允许一次添加 1 行）。
* `range.getColumnsAfter(count: number)` 和 `range.getColumnsBefore(count: number)` 分别用于获取当前 Range 对象的右/左侧的一定数量的列。
* 获取项或 NULL 对象函数：此功能允许使用键获取对象。如果没有对象，返回的对象的 isNullObject 属性为 true。这样一来，开发者可以检查对象是否存在，而无需通过异常处理来进行处理。适用于工作表、已命名项、绑定、图表系列等

    ```javascript
    worksheet.GetItemOrNullObject()
    ```

|对象| 最近更新| 说明|要求集|
|:----|:----|:----|:----|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_方法_> getCount()|获取集合中的绑定数量。|1.4|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_方法_> getItemOrNullObject(id: string)|按 ID 获取 Binding 对象。如果没有 Binding 对象，将返回 NULL 对象。|1.4|
|[chartCollection](/javascript/api/excel/excel.chartcollection)|_方法_> getCount()|返回工作表中的图表数。|1.4|
|[chartCollection](/javascript/api/excel/excel.chartcollection)|_方法_> getItemOrNullObject(name: string)|使用图表名称获取图表。如果存在多个名称相同的图表，将返回第一个图表。|1.4|
|[chartPointsCollection](/javascript/api/excel/excel.chartpointscollection)|_方法_> getCount()|返回系列中的图表点数。|1.4|
|[chartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|_方法_> getCount()|返回集合中的系列数量。|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_属性_ > comment|表示与此名称相关联的注释。|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_属性_ > scope|指明是否将 name 限定到工作簿或特定工作表。只读。可取值为：Equal、Greater、GreaterEqual、Less、LessEqual、NotEqual。|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_关系_ > worksheet|返回已命名项限定到的工作表。如果项改为限定到工作簿，将引发错误。只读。|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_关系_ > worksheetOrNullObject|返回已命名项限定到的工作表。如果项改为限定到工作簿，将返回 NULL 对象。只读。|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_方法_> delete()|删除给定的名称。|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_方法_> getRangeOrNullObject()|返回与名称相关联的 Range 对象。如果已命名项的类型不是 Range，将返回 NULL 对象。|1.4|
|[namedItemCollection](/javascript/api/excel/excel.nameditemcollection)|_方法_> 添加 (名称： 字符串，参考： 范围或字符串，注释： 字符串)|将新名称添加到给定范围的集合。|1.4|
|[namedItemCollection](/javascript/api/excel/excel.nameditemcollection)|_方法_> addFormulaLocal (名称： 字符串，公式： 字符串，注释： 字符串)|使用用户的公式区域设置，将新名称添加到给定范围的集合。|1.4|
|[namedItemCollection](/javascript/api/excel/excel.nameditemcollection)|_方法_> getCount()|获取集合中已命名项的数量。|1.4|
|[namedItemCollection](/javascript/api/excel/excel.nameditemcollection)|_方法_> getItemOrNullObject(name: string)|按 NamedItem 对象的名称获取此对象。如果没有 NamedItem 对象，将返回 NULL 对象。|1.4|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|_方法_> getCount()|获取集合中的数据透视表的数量。|1.4|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|_方法_> getItemOrNullObject(name: string)|按 PivotTable 对象的名称获取此对象。如果没有 PivotTable 对象，将返回 NULL 对象。|1.4|
|[range](/javascript/api/excel/excel.range)|_方法_> getIntersectionOrNullObject (anotherRange： 范围或字符串)|获取表示指定区域的矩形交集的 range 对象。如果找不到任何交集，则此方法返回空对象。|1.4|
|[range](/javascript/api/excel/excel.range)|_方法_> getUsedRangeOrNullObject(valuesOnly: bool)|返回指定 Range 对象的所用区域。如果区域内没有使用单元格，此函数将返回 NULL 对象。|1.4|
|[rangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|_方法_> getCount()|获取集合中 RangeView 对象的数量。|1.4|
|[setting](/javascript/api/excel/excel.setting)|_属性_ > key|返回表示 setting 对象的 ID 的键。只读。|1.4|
|[setting](/javascript/api/excel/excel.setting)|_属性_ > value|表示为此设置存储的值。|1.4|
|[setting](/javascript/api/excel/excel.setting)|_方法_> delete()|删除 Setting 对象。|1.4|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_属性_ > items|一组 setting 对象。只读。|1.4|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_方法_> 添加 (密钥： 字符串，值: （任何）)|设置指定的 Setting 对象，或将其添加到工作簿中。|1.4|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_方法_> getCount()|获取集合中的 Setting 对象的数量。|1.4|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_方法_> getItem(key: string)|按键获取 Setting 项。|1.4|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_方法_> getItemOrNullObject(key: string)|按键获取 Setting 项。如果没有 Setting 项，将返回 NULL 对象。|1.4|
|[settingsChangedEventArgs](/javascript/api/excel/excel.settingschangedeventargs)|_关系_ > settings|获取表示引发了 SettingsChanged 事件的绑定的 Setting 对象。|1.4|
|[tableCollection](/javascript/api/excel/excel.tablecollection)|_方法_> getCount()]|获取集合中的表数量。|1.4|
|[tableCollection](/javascript/api/excel/excel.tablecollection)|_方法_> getItemOrNullObject (密钥： 数字或字符串)|按名称或 ID 获取表。如果没有表，将返回 NULL 对象。|1.4|
|[tableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|_方法_> getCount()|获取表中的列数。|1.4|
|[tableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|_方法_> getItemOrNullObject (密钥： 数字或字符串)|按名称或 ID 获取 column 对象。如果没有 column 对象，将返回 NULL 对象。|1.4|
|[tableRowCollection](/javascript/api/excel/excel.tablerowcollection)|_方法_> getCount()|获取表格中的行数。|1.4|
|[workbook](/javascript/api/excel/excel.workbook)|_关系_ > settings|表示一组与 workbook 相关联的 setting 对象。只读。|1.4|
|[worksheet](/javascript/api/excel/excel.worksheet)|_关系_ > names|一组范围限定到当前工作表的名称。只读。|1.4|
|[worksheet](/javascript/api/excel/excel.worksheet)|_方法_> getUsedRangeOrNullObject(valuesOnly: bool)|使用的区域是包含分配了值或格式的任意单元格的最小区域。如果整个工作表为空，此函数将返回 NULL 对象。|1.4|
|[worksheetCollection](/javascript/api/excel/excel.worksheetcollection)|_方法_> getCount(visibleOnly: bool)|获取集合中的工作表数量。|1.4|
|[worksheetCollection](/javascript/api/excel/excel.worksheetcollection)|_方法_> getItemOrNullObject(key: string)|按 Worksheet 对象的名称或 ID 获取此对象。如果没有 Worksheet 对象，将返回 NULL 对象。|1.4|

## <a name="whats-new-in-excel-javascript-api-13"></a>Excel JavaScript API 1.3 的最近更新

下面介绍了要求集 1.3 中 Excel JavaScript API 的新增内容。

|对象| 最近更新| 说明|要求集|
|:----|:----|:----|:----|
|[binding](/javascript/api/excel/excel.binding)|_方法_> delete()|删除 binding 对象。|1.3|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_方法_> 添加 (范围： 范围或字符串，bindingType: string、 id： 字符串)|将新的 binding 对象添加到特定区域。|1.3|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_方法_> addFromNamedItem (名称： string、 bindingType: string、 id： 字符串)|根据工作簿中的命名项添加新的 binding 对象。|1.3|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_方法_> addFromSelection (bindingType: string、 id： 字符串)|根据当前选择的内容添加新的 binding 对象。|1.3|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_方法_> getItemOrNull(id: string)|按 ID 获取 binding 对象。如果 binding 对象不存在，则返回对象的 isNull 属性为 true。|1.3|
|[chartCollection](/javascript/api/excel/excel.chartcollection)|_方法_> getItemOrNull(name: string)|使用图表名称获取图表。如果存在多个名称相同的图表，将返回第一个图表。|1.3|
|[namedItemCollection](/javascript/api/excel/excel.nameditemcollection)|_方法_> getItemOrNull(name: string)|按 nameditem 对象的名称获取此对象。如果 nameditem 对象不存在，则返回对象的 isNull 属性为 true。|1.3|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_属性_ > name|PivotTable 对象的名称。|1.3|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_关系_ > worksheet|包含当前 PivotTable 对象的工作表。只读。|1.3|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_方法_> refresh()|刷新 PivotTable 对象。|1.3|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|_属性_ > items|一组 PivotTable 对象。只读。|1.3|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|_方法_> getItem(name: string)|按名称获取 PivotTable 对象。|1.3|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|_方法_> getItemOrNull(name: string)|按名称获取 PivotTable 对象。如果 PivotTable 对象不存在，则返回对象的 isNull 属性为 true。|1.3|
|[range](/javascript/api/excel/excel.range)|_方法_> getIntersectionOrNull (anotherRange： 范围或字符串)|获取表示指定区域的矩形交集的 range 对象。如果找不到任何交集，则此方法返回空对象。|1.3|
|[range](/javascript/api/excel/excel.range)|_方法_> getVisibleView()|表示当前 range 对象的可见行。|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_属性_ > cellAddresses|表示 RangeView 的单元格地址。只读。|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_属性_ > columnCount|返回可见列数。只读。|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_属性_ > formulas|表示采用 A1 表示法的公式。|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_属性_ > formulasLocal|使用用户语言和数字格式区域设置表示采用 A1 表示法的公式。例如，用英语表示的公式 "=SUM(A1, introduced in 1.5)" 在德语中将变为 "=SUMME(A1; 1,5)"。|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_属性_ > formulasR1C1|表示采用 R1C1 表示法的公式。|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_属性_ > index|返回表示 RangeView 的索引的值。只读。|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_属性_ > numberFormat|表示 Excel 中指定单元格的数字格式代码。|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_属性_ > rowCount|返回可见行数。只读。|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_属性_ > text|指定区域的文本值。文本值与单元格宽度无关。在 Excel UI 中替代 # 符号不会影响 API 返回的文本值。只读。|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_属性_ > valueTypes|表示每个单元格的数据类型。只读。可能的值是：Unknown、Empty、String、Integer、Double、Boolean、Error。|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_属性_ > values|表示指定的 RangeView 的原始值。返回的数据可能是字符串、数字，也可能是布尔值。包含错误的单元格将返回错误字符串。|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_关系_ > rows|表示一组与 range 相关联的 RangeView。只读。|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_方法_> getRange()|获取与当前 RangeView 相关联的父 range。|1.3|
|[rangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|_属性_ > items|一组 rangeView 对象。只读。|1.3|
|[rangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|_方法_> getItemAt(index: number)|按索引获取 RangeView 行。从零开始编制索引。|1.3|
|[setting](/javascript/api/excel/excel.setting)|_属性_ > key|返回表示 setting 对象的 ID 的键。只读。|1.3|
|[setting](/javascript/api/excel/excel.setting)|_方法_> delete()|删除 setting 对象。|1.3|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_属性_ > items|一组 setting 对象。只读。|1.3|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_方法_> getItem(key: string)|按键获取 setting 项。|1.3|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_方法_> getItemOrNull(key: string)|按键获取 setting 项。如果 setting 对象不存在，则返回对象的 isNull 属性为 true。|1.3|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_方法_> 设置 (密钥： 字符串，值： 字符串)|设置指定的 setting 对象，或将其添加到工作簿中。|1.3|
|[settingsChangedEventArgs](/javascript/api/excel/excel.settingschangedeventargs)|_关系_ > settingCollection|获取表示引发了 SettingsChanged 事件的 binding 的 setting 对象。|1.3|
|[table](/javascript/api/excel/excel.table)|_属性_ > highlightFirstColumn|指明第一列是否包含特殊格式。|1.3|
|[table](/javascript/api/excel/excel.table)|_属性_ > highlightLastColumn|指明最后一列是否包含特殊格式。|1.3|
|[table](/javascript/api/excel/excel.table)|_属性_ > showBandedColumns|指明列是否采用镶边格式来以不同的方式突出显示奇数列与偶数列，让表更易于阅读。|1.3|
|[table](/javascript/api/excel/excel.table)|_属性_ > showBandedRows|指明行是否采用镶边格式来以不同的方式突出显示奇数行与偶数行，让表更易于阅读。|1.3|
|[table](/javascript/api/excel/excel.table)|_属性_ > showFilterButton|指明是否在每个列标题的顶部显示筛选器按钮。仅当 table 中包含标题行时，才允许设定此设置。|1.3|
|[tableCollection](/javascript/api/excel/excel.tablecollection)|_方法_> getItemOrNull (密钥： 数字或字符串)|按名称或 ID 获取 table 对象。如果 table 对象不存在，则返回对象的 isNull 属性为 true。|1.3|
|[tableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|_方法_> getItemOrNull (密钥： 数字或字符串)|按名称或 ID 获取 column 对象。如果 column 对象不存在，则返回对象的 isNull 属性为 true。|1.3|
|[workbook](/javascript/api/excel/excel.workbook)|_关系_ > pivotTables|表示一组与 workbook 相关联的 PivotTable 对象。只读。|1.3|
|[workbook](/javascript/api/excel/excel.workbook)|_关系_ > settings|表示一组与 workbook 相关联的 setting 对象。只读。|1.3|
|[worksheet](/javascript/api/excel/excel.worksheet)|_关系_ > pivotTables|一组属于 worksheet 的 PivotTable 对象。只读。|1.3|

## <a name="whats-new-in-excel-javascript-api-12"></a>Excel JavaScript API 1.2 的最近更新

下面介绍了要求集 1.2 中 Excel JavaScript API 的新增内容。

|对象| 最近更新| 说明|要求集|
|:----|:----|:----|:----|
|[chart](/javascript/api/excel/excel.chart)|_属性_ > id|根据其在集合中的位置获取图表。只读。|1.2|
|[chart](/javascript/api/excel/excel.chart)|_关系_ > worksheet|包含当前 chart 的 worksheet 对象。只读。|1.2|
|[chart](/javascript/api/excel/excel.chart)|_方法_> getImage (高度： 数字、 宽度： 号码后，fittingMode： 字符串)|通过缩放 chart 以适应指定的尺寸，将 chart 呈现为 base64 编码的图像。|1.2|
|[filter](/javascript/api/excel/excel.filter)|_关系_ > criteria|给定列上当前应用的筛选器。只读。|1.2|
|[filter](/javascript/api/excel/excel.filter)|_方法_> apply(criteria: FilterCriteria)|在给定列中应用给定的筛选条件。|1.2|
|[filter](/javascript/api/excel/excel.filter)|_方法_> applyBottomItemsFilter(count: number)|将“Bottom Item”筛选器应用于列，以获取给定数量的元素。|1.2|
|[filter](/javascript/api/excel/excel.filter)|_方法_> applyBottomPercentFilter(percent: number)]|将“Bottom Percent”筛选器应用于列，以获取给定比例的元素。|1.2|
|[filter](/javascript/api/excel/excel.filter)|_方法_> applyCellColorFilter(color: string)|将“Cell Color”筛选器应用于列，以获取给定颜色。|1.2|
|[filter](/javascript/api/excel/excel.filter)|_方法_> applyCustomFilter (criteria1: string、 criteria2: string、 运算符： 字符串)|将“Icon”筛选器应用于列，以获取给定的条件字符串。|1.2|
|[filter](/javascript/api/excel/excel.filter)|_方法_> applyDynamicFilter(criteria: string)|将“Dynamic”筛选器应用于列。|1.2|
|[filter](/javascript/api/excel/excel.filter)|_方法_> applyFontColorFilter(color: string)|将“Font Color”筛选器应用于列，以获取给定颜色。|1.2|
|[filter](/javascript/api/excel/excel.filter)|_方法_> applyIconFilter(icon: Icon)|将“Icon”筛选器应用于列，以获取给定 icon。|1.2|
|[filter](/javascript/api/excel/excel.filter)|_方法_> applyTopItemsFilter(count: number)|将“Top Item”筛选器应用于列，以获取给定数量的元素。|1.2|
|[filter](/javascript/api/excel/excel.filter)|_方法_> applyTopPercentFilter(percent: number)|将“Top Percent”筛选器应用于列，以获取给定比例的元素。|1.2|
|[filter](/javascript/api/excel/excel.filter)|_方法_> applyValuesFilter (值: （)）|将“Values”筛选器应用于列，以获取给定值。|1.2|
|[filter](/javascript/api/excel/excel.filter)|_方法_> clear()|清除给定列上的 filter。|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_属性_ > color|用于筛选单元格的 HTML 颜色字符串。与“cellColor”和“fontColor”筛选一起使用。|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_属性_ > criterion1|第一个条件用于筛选数据。在“自定义”筛选中用作运算符。|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_属性_ > criterion2|第二个条件用于筛选数据。在“自定义”筛选中仅用作运算符。|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_属性_ > dynamicCriteria|Excel.DynamicFilterCriteria 集中的动态条件将应用于此列。与“动态”筛选一起使用。可能的值是：Unknown、AboveAverage、AllDatesInPeriodApril、AllDatesInPeriodAugust、AllDatesInPeriodDecember、AllDatesInPeriodFebruray、AllDatesInPeriodJanuary、AllDatesInPeriodJuly、AllDatesInPeriodJune、AllDatesInPeriodMarch、AllDatesInPeriodMay、AllDatesInPeriodNovember、AllDatesInPeriodOctober、AllDatesInPeriodQuarter1、AllDatesInPeriodQuarter2、AllDatesInPeriodQuarter3、AllDatesInPeriodQuarter4、AllDatesInPeriodSeptember、BelowAverage、LastMonth、LastQuarter、LastWeek、LastYear、NextMonth、NextQuarter、NextWeek、NextYear、ThisMonth、ThisQuarter、ThisWeek、ThisYear、Today、Tomorrow、YearToDate、Yesterday。|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_属性_ > filterOn|filter 使用的属性，用于确定值是否应一直可见。可取值为：BottomItems、BottomPercent、CellColor、Dynamic、FontColor、Values、TopItems、TopPercent、Icon、Custom。|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_属性_ > operator|使用“自定义”筛选器时，用于组合条件 1 和 2 的运算符。可取值为：And、Or。|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_属性_ > values|一组用于“values”筛选器的值。|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_关系_ > icon|用于筛选单元格的图标。与“图标”筛选一起使用。|1.2|
|[filterDatetime](/javascript/api/excel/excel.filterdatetime)|_属性_ > date|用于筛选数据的采用 ISO8601 格式的日期。|1.2|
|[filterDatetime](/javascript/api/excel/excel.filterdatetime)|_属性_ > specificity|用于保留数据的日期的具体程度。例如，如果日期是 2005-04-02 并且将特殊性设置为“月”，则筛选操作将保留包含 2009 年 4 月日期的所有行。可能的值是：Year、Monday、Day、Hour、Minute、Second。|1.2|
|[formatProtection](/javascript/api/excel/excel.formatprotection)|_属性_ > formulaHidden|表示 Excel 是否隐藏区域中的单元格公式。指示整个区域不具有统一公式隐藏设置的空值。|1.2|
|[formatProtection](/javascript/api/excel/excel.formatprotection)|_属性_ > locked|指示 Excel 是否锁定对象中的单元格。指示整个区域不具有统一锁定设置的空值。|1.2|
|[icon](/javascript/api/excel/excel.icon)|_属性_ > index|表示 icon 在给定集中的索引。|1.2|
|[icon](/javascript/api/excel/excel.icon)|_属性_ > set|表示图标所属的集合。可能的值是：Invalid、ThreeArrows、ThreeArrowsGray、ThreeFlags、ThreeTrafficLights1、ThreeTrafficLights2、ThreeSigns、ThreeSymbols、ThreeSymbols2、FourArrows、FourArrowsGray、FourRedToBlack、FourRating、FourTrafficLights、FiveArrows、FiveArrowsGray、FiveRating、FiveQuarters、ThreeStars、ThreeTriangles、FiveBoxes。|1.2|
|[range](/javascript/api/excel/excel.range)|_属性_ > columnHidden|表示当前 range 的所有列均已隐藏。|1.2|
|[range](/javascript/api/excel/excel.range)|_属性_ > formulasR1C1|表示采用 R1C1 表示法的公式。|1.2|
|[range](/javascript/api/excel/excel.range)|_属性_ > hidden|表示当前区域中的所有单元格是否隐藏。只读。|1.2|
|[range](/javascript/api/excel/excel.range)|_属性_ > rowHidden|表示当前 range 的所有行均已隐藏。|1.2|
|[range](/javascript/api/excel/excel.range)|_关系_ > sort|表示当前 range 的区域排序。只读。|1.2|
|[range](/javascript/api/excel/excel.range)|_方法_> merge(across: bool)|将 range 单元格合并到 worksheet 的一个区域内。|1.2|
|[range](/javascript/api/excel/excel.range)|_方法_> unmerge()|将 range 单元格拆分为单个单元格。|1.2|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_属性_ > columnWidth|获取或设置区域内的所有列的宽度。如果列宽不统一，则返回 NULL。|1.2|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_属性_ > rowHeight|获取或设置区域中所有行的高度。如果行高不统一，则返回 NULL。|1.2|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_关系_ > protection|返回某一区域的格式保护对象。只读。|1.2|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_方法_> autofitColumns()|根据列中的当前数据，更改当前 range 内所有列的宽度，以达到最佳显示效果。|1.2|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_方法_> autofitRows()|根据列中的当前数据，更改当前 range 内所有行的高度，以达到最佳显示效果。|1.2|
|[rangeReference](/javascript/api/excel/excel.rangereference)|_属性_ > address|表示当前 range 对象的可见行。|1.2|
|[rangeSort](/javascript/api/excel/excel.rangesort)|_方法_> 应用 (字段： SortField，matchCase: bool，hasHeaders: bool，方向： 字符串，方法： 字符串)|执行排序操作。|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_属性_ > ascending|表示是否执行升序排序。|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_属性_ > color|表示按字体或单元格颜色进行排序时，条件的目标颜色。|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_属性_ > dataOption|表示此字段的其他排序选项。可能的值是：Normal、TextAsNumber。|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_属性_ > key|表示条件所在的列（或行，具体取决于排序方向）。表示与第一列（或行）的偏移量。|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_属性_ > sortOn|表示此条件的排序类型。可能的值是：Value、CellColor、FontColor、Icon。|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_关系_ > icon|表示对单元格图标进行排序时，条件的目标图标。|1.2|
|[table](/javascript/api/excel/excel.table)|_关系_ > sort|表示表的排序。只读。|1.2|
|[table](/javascript/api/excel/excel.table)|_关系_ > worksheet|包含当前表的工作表。只读。|1.2|
|[table](/javascript/api/excel/excel.table)|_方法_> clearFilters()|清除当前在 table 中应用的所有 filter。|1.2|
|[table](/javascript/api/excel/excel.table)|_方法_> convertToRange()|将表转换为普通单元格区域。保留所有数据。|1.2|
|[table](/javascript/api/excel/excel.table)|_方法_> reapplyFilters()|重新应用当前在 table 上应用的所有 filter。|1.2|
|[tableColumn](/javascript/api/excel/excel.tablecolumn)|_关系_ > filter|检索应用于列的筛选器。只读。|1.2|
|[tableSort](/javascript/api/excel/excel.tablesort)|_属性_ > matchCase|表示最后一次对表进行排序时大小写是否有影响。只读。|1.2|
|[tableSort](/javascript/api/excel/excel.tablesort)|_属性_ > method|表示最后一次对表排序所使用的中文字符排序方法。只读。可能的值是：PinYin、StrokeCount。|1.2|
|[tableSort](/javascript/api/excel/excel.tablesort)|_关系_ > fields|表示最后一次对表排序所使用的当前条件。只读。|1.2|
|[tableSort](/javascript/api/excel/excel.tablesort)|_方法_> 应用 (字段： SortField，matchCase: bool，方法： 字符串)|执行排序操作。|1.2|
|[tableSort](/javascript/api/excel/excel.tablesort)|_方法_> clear()|清除表上的当前排序。尽管这不能修改表的排序，但它会清除标题按钮的状态。|1.2|
|[tableSort](/javascript/api/excel/excel.tablesort)|_方法_> reapply()|对 table 重新应用当前的排序参数。|1.2|
|[workbook](/javascript/api/excel/excel.workbook)|_关系_ > functions|表示包含此工作簿的 Excel 应用程序实例。只读。|1.2|
|[worksheet](/javascript/api/excel/excel.worksheet)|_关系_ > protection|返回表工作表的工作表保护对象。只读。|1.2|
|[worksheetProtection](/javascript/api/excel/excel.worksheetprotection)|_属性_ > protected|指明 worksheet 是否受保护。只读。只读。|1.2|
|[worksheetProtection](/javascript/api/excel/excel.worksheetprotection)|_关系_ > options|工作表保护选项。只读。|1.2|
|[worksheetProtection](/javascript/api/excel/excel.worksheetprotection)|_方法_> protect(options: WorksheetProtectionOptions)|保护 worksheet。如果 worksheet 处于受保护状态，则无法执行此方法。|1.2|
|[worksheetProtection](/javascript/api/excel/excel.worksheetprotection)|_方法_> unprotect()|解除对 worksheet 的保护。|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_属性_ > allowAutoFilter|表示允许使用自动筛选功能的工作表保护选项。|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_属性_ > allowDeleteColumns|表示允许删除列的工作表保护选项。|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_属性_ > allowDeleteRows|表示允许删除行的工作表保护选项。|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_属性_ > allowFormatCells|表示允许格式化单元格的工作表保护选项。|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_属性_ > allowFormatColumns|表示允许格式化列的工作表保护选项。|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_属性_ > allowFormatRows|表示允许格式化行的工作表保护选项。|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_属性_ > allowInsertColumns|表示允许插入列的工作表保护选项。|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_属性_ > allowInsertHyperlinks|表示允许插入超链接的工作表保护选项。|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_属性_ > allowInsertRows|表示允许插入行的工作表保护选项。|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_属性_ > allowPivotTables|表示允许使用数据透视表功能的工作表保护选项。|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_属性_ > allowSort|表示允许使用排序功能的工作表保护选项。|1.2|

## <a name="excel-javascript-api-11"></a>Excel JavaScript API 1.1

Excel JavaScript API 1.1 是第一个版本的 api。 有关 API 的详细信息，请参阅[Excel 的 JavaScript API](/javascript/api/excel)参考主题。

## <a name="see-also"></a>另请参阅

- [Office 版本和要求集](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [指定 Office 主机和 API 要求](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Office 外接程序 XML 清单](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests)
