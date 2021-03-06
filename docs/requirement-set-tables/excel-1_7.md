| 类 | 域 | 说明 |
|:---|:---|:---|
|[Chart](/javascript/api/excel/excel.chart)|[chartType](/javascript/api/excel/excel.chart#charttype)|指定图表的类型。|
||[id](/javascript/api/excel/excel.chart#id)|图表的唯一 ID。|
||[showAllFieldButtons](/javascript/api/excel/excel.chart#showallfieldbuttons)|指定是否在数据透视图上显示所有字段按钮。|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[border](/javascript/api/excel/excel.chartareaformat#border)|表示图表区的边框格式，包括颜色、线条和粗细。|
|[ChartAxes](/javascript/api/excel/excel.chartaxes)|[getItem (类型：Excel.ChartAxisType， group？： Excel.ChartAxisGroup) ](/javascript/api/excel/excel.chartaxes#getitem-type--group-)|返回通过类型和组标识的特定轴。|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[baseTimeUnit](/javascript/api/excel/excel.chartaxis#basetimeunit)|指定指定分类轴的基本单位。|
||[categoryType](/javascript/api/excel/excel.chartaxis#categorytype)|指定分类轴类型。|
||[displayUnit](/javascript/api/excel/excel.chartaxis#displayunit)|表示轴显示单位。|
||[logBase](/javascript/api/excel/excel.chartaxis#logbase)|指定使用对数刻度时对数的底数。|
||[majorTickMark](/javascript/api/excel/excel.chartaxis#majortickmark)|指定指定坐标轴的主要刻度线的类型。|
||[majorTimeUnitScale](/javascript/api/excel/excel.chartaxis#majortimeunitscale)|指定 CategoryType 属性设置为 TimeScale 时分类轴的主要单位刻度值。|
||[minorTickMark](/javascript/api/excel/excel.chartaxis#minortickmark)|指定指定坐标轴的次要刻度线的类型。|
||[minorTimeUnitScale](/javascript/api/excel/excel.chartaxis#minortimeunitscale)|指定 CategoryType 属性设置为 TimeScale 时分类轴的次要单位刻度值。|
||[axisGroup](/javascript/api/excel/excel.chartaxis#axisgroup)|指定指定坐标轴的组。|
||[customDisplayUnit](/javascript/api/excel/excel.chartaxis#customdisplayunit)|指定自定义轴显示单位值。|
||[height](/javascript/api/excel/excel.chartaxis#height)|指定图表轴的高度（以点表示）。|
||[left](/javascript/api/excel/excel.chartaxis#left)|指定从坐标轴左边缘到图表区左侧的距离（以点表示）。|
||[top](/javascript/api/excel/excel.chartaxis#top)|指定从坐标轴上边缘到图表区顶部的距离（以点表示）。|
||[type](/javascript/api/excel/excel.chartaxis#type)|指定坐标轴类型。|
||[width](/javascript/api/excel/excel.chartaxis#width)|指定图表轴的宽度（以点表示）。|
||[reversePlotOrder](/javascript/api/excel/excel.chartaxis#reverseplotorder)|指定 Excel 是否绘制从最后一个数据点到第一个数据点。|
||[scaleType](/javascript/api/excel/excel.chartaxis#scaletype)|指定数值轴刻度类型。|
||[setCategoryNames (sourceData：Range) ](/javascript/api/excel/excel.chartaxis#setcategorynames-sourcedata-)|设置指定轴的所有分类名称。|
||[setCustomDisplayUnit (值：number) ](/javascript/api/excel/excel.chartaxis#setcustomdisplayunit-value-)|将轴显示单位设为自定义值。|
||[showDisplayUnitLabel](/javascript/api/excel/excel.chartaxis#showdisplayunitlabel)|指定坐标轴显示单位标签是否可见。|
||[tickLabelPosition](/javascript/api/excel/excel.chartaxis#ticklabelposition)|在指定坐标轴上指定刻度线标签的位置。|
||[tickLabelSpacing](/javascript/api/excel/excel.chartaxis#ticklabelspacing)|指定刻度线标签之间的分类数或系列数。|
||[tickMarkSpacing](/javascript/api/excel/excel.chartaxis#tickmarkspacing)|指定刻度线之间的分类数或系列数。|
||[visible](/javascript/api/excel/excel.chartaxis#visible)|指定坐标轴是否可见。|
|[ChartBorder](/javascript/api/excel/excel.chartborder)|[color](/javascript/api/excel/excel.chartborder#color)|表示图表中的边框颜色的 HTML 颜色代码。|
||[lineStyle](/javascript/api/excel/excel.chartborder#linestyle)|表示边框的线条样式。|
||[weight](/javascript/api/excel/excel.chartborder#weight)|表示边框的粗细，以磅为单位。|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[position](/javascript/api/excel/excel.chartdatalabel#position)|表示数据标签的位置的 DataLabelPosition 值。|
||[分隔符](/javascript/api/excel/excel.chartdatalabel#separator)|该字符串表示用于图表中数据标签的分隔符。|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabel#showbubblesize)|指定数据标签气泡大小是否可见。|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabel#showcategoryname)|指定数据标签类别名称是否可见。|
||[showLegendKey](/javascript/api/excel/excel.chartdatalabel#showlegendkey)|指定数据标签图例项标示是否可见。|
||[showPercentage](/javascript/api/excel/excel.chartdatalabel#showpercentage)|指定数据标签百分比是否可见。|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabel#showseriesname)|指定数据标签系列名称是否可见。|
||[showValue](/javascript/api/excel/excel.chartdatalabel#showvalue)|指定数据标签值是否可见。|
|[ChartFormatString](/javascript/api/excel/excel.chartformatstring)|[font](/javascript/api/excel/excel.chartformatstring#font)|表示字体属性，如字体名称、字号、颜色等。|
|[ChartLegend](/javascript/api/excel/excel.chartlegend)|[height](/javascript/api/excel/excel.chartlegend#height)|指定图表上图例的高度（以点表示）。|
||[left](/javascript/api/excel/excel.chartlegend#left)|指定图表上图例的左侧（以点表示）。|
||[legendEntries](/javascript/api/excel/excel.chartlegend#legendentries)|表示图例中 legendEntries 的集合。|
||[showShadow](/javascript/api/excel/excel.chartlegend#showshadow)|指定图例在图表上是否具有阴影。|
||[top](/javascript/api/excel/excel.chartlegend#top)|指定图表图例的顶部。|
||[width](/javascript/api/excel/excel.chartlegend#width)|指定图表上的图例的宽度（以点表示）。|
|[ChartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|[visible](/javascript/api/excel/excel.chartlegendentry#visible)|表示图表图例条目可见。|
|[ChartLegendEntryCollection](/javascript/api/excel/excel.chartlegendentrycollection)|[getCount()](/javascript/api/excel/excel.chartlegendentrycollection#getcount--)|返回集合中的 legendEntry 数量。|
||[getItemAt(index: number)](/javascript/api/excel/excel.chartlegendentrycollection#getitemat-index-)|返回给定索引处的 legendEntry。|
||[items](/javascript/api/excel/excel.chartlegendentrycollection#items)|获取此集合中已加载的子项。|
|[ChartLineFormat](/javascript/api/excel/excel.chartlineformat)|[lineStyle](/javascript/api/excel/excel.chartlineformat#linestyle)|代表线条样式。|
||[weight](/javascript/api/excel/excel.chartlineformat#weight)|表示线条的粗细（以磅为单位）。|
|[ChartPoint](/javascript/api/excel/excel.chartpoint)|[hasDataLabel](/javascript/api/excel/excel.chartpoint#hasdatalabel)|表示一个数据点是否具有数据标签。|
||[markerBackgroundColor](/javascript/api/excel/excel.chartpoint#markerbackgroundcolor)|数据点的标记背景色的 HTML 颜色代码表示 (，例如，#FF0000表示红色) 。|
||[markerForegroundColor](/javascript/api/excel/excel.chartpoint#markerforegroundcolor)|数据点的标记前景色的 HTML 颜色代码表示 (，例如，#FF0000表示红色) 。|
||[markerSize](/javascript/api/excel/excel.chartpoint#markersize)|表示数据点的标记大小。|
||[markerStyle](/javascript/api/excel/excel.chartpoint#markerstyle)|表示图表数据点的标记样式。|
||[dataLabel](/javascript/api/excel/excel.chartpoint#datalabel)|返回图表点的数据标签。|
|[ChartPointFormat](/javascript/api/excel/excel.chartpointformat)|[border](/javascript/api/excel/excel.chartpointformat#border)|表示图表数据点的边框格式，包括颜色、样式和粗细信息。|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[chartType](/javascript/api/excel/excel.chartseries#charttype)|表示系列的图表类型。|
||[delete()](/javascript/api/excel/excel.chartseries#delete--)|删除 chart series 对象。|
||[doughnutHoleSize](/javascript/api/excel/excel.chartseries#doughnutholesize)|表示图表系列的圆环孔大小。|
||[已筛选](/javascript/api/excel/excel.chartseries#filtered)|指定是否筛选系列。|
||[gapWidth](/javascript/api/excel/excel.chartseries#gapwidth)|表示图表系列的间隙宽度。|
||[hasDataLabels](/javascript/api/excel/excel.chartseries#hasdatalabels)|指定系列是否具有数据标签。|
||[markerBackgroundColor](/javascript/api/excel/excel.chartseries#markerbackgroundcolor)|指定图表系列的标记背景色。|
||[markerForegroundColor](/javascript/api/excel/excel.chartseries#markerforegroundcolor)|指定图表系列的标记前景色。|
||[markerSize](/javascript/api/excel/excel.chartseries#markersize)|指定图表系列的标记大小。|
||[markerStyle](/javascript/api/excel/excel.chartseries#markerstyle)|指定图表系列的标记样式。|
||[plotOrder](/javascript/api/excel/excel.chartseries#plotorder)|指定图表组中图表系列的绘制顺序。|
||[trendlines](/javascript/api/excel/excel.chartseries#trendlines)|系列中趋势线的集合。|
||[setBubbleSizes (sourceData：Range) ](/javascript/api/excel/excel.chartseries#setbubblesizes-sourcedata-)|设置图表系列的气泡大小。|
||[setValues (sourceData：Range) ](/javascript/api/excel/excel.chartseries#setvalues-sourcedata-)|设置图表系列的值。|
||[setXAxisValues (sourceData：Range) ](/javascript/api/excel/excel.chartseries#setxaxisvalues-sourcedata-)|设置图表系列的 X 轴的值。|
||[showShadow](/javascript/api/excel/excel.chartseries#showshadow)|指定系列是否具有阴影。|
||[平滑](/javascript/api/excel/excel.chartseries#smooth)|指定系列是否平滑。|
|[ChartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|[添加 (名称？： string， index？： number) ](/javascript/api/excel/excel.chartseriescollection#add-name--index-)|向集合添加新系列。|
|[ChartTitle](/javascript/api/excel/excel.charttitle)|[getSubstring (start： number， length： number) ](/javascript/api/excel/excel.charttitle#getsubstring-start--length-)|获取图表标题的子字符串。|
||[horizontalAlignment](/javascript/api/excel/excel.charttitle#horizontalalignment)|指定图表标题的水平对齐方式。|
||[left](/javascript/api/excel/excel.charttitle#left)|指定从图表标题的左边缘到图表区左边缘的距离（以点表示）。|
||[position](/javascript/api/excel/excel.charttitle#position)|表示图表标题的位置。|
||[height](/javascript/api/excel/excel.charttitle#height)|返回图表标题的高度，以磅为单位。|
||[width](/javascript/api/excel/excel.charttitle#width)|指定图表标题的宽度（以点表示）。|
||[setFormula (公式：字符串) ](/javascript/api/excel/excel.charttitle#setformula-formula-)|设置一个字符串值，用于表示采用 A1 表示法的图表标题的公式。|
||[showShadow](/javascript/api/excel/excel.charttitle#showshadow)|表示一个布尔值，用于确定图表标题是否具有阴影。|
||[textOrientation](/javascript/api/excel/excel.charttitle#textorientation)|指定文本面向图表标题的角度。|
||[top](/javascript/api/excel/excel.charttitle#top)|指定从图表标题的上边缘到图表区域顶部的距离（以点表示）。|
||[verticalAlignment](/javascript/api/excel/excel.charttitle#verticalalignment)|指定图表标题的垂直对齐方式。|
|[ChartTitleFormat](/javascript/api/excel/excel.charttitleformat)|[border](/javascript/api/excel/excel.charttitleformat#border)|表示图表标题的边框格式，包括颜色、线条tyle 和粗细。|
|[ChartTrendline](/javascript/api/excel/excel.charttrendline)|[delete()](/javascript/api/excel/excel.charttrendline#delete--)|删除 Trendline 对象。|
||[intercept](/javascript/api/excel/excel.charttrendline#intercept)|表示趋势线的截距值。|
||[movingAveragePeriod](/javascript/api/excel/excel.charttrendline#movingaverageperiod)|表示图表趋势线的周期。|
||[name](/javascript/api/excel/excel.charttrendline#name)|表示趋势线的名称。|
||[polynomialOrder](/javascript/api/excel/excel.charttrendline#polynomialorder)|表示图表趋势线的顺序。|
||[format](/javascript/api/excel/excel.charttrendline#format)|表示图表趋势线的格式。|
||[type](/javascript/api/excel/excel.charttrendline#type)|表示图表趋势线的类型。|
|[ChartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|[添加 (类型？：Excel.ChartTrendlineType) ](/javascript/api/excel/excel.charttrendlinecollection#add-type-)|向趋势线集合添加新的趋势线。|
||[getCount()](/javascript/api/excel/excel.charttrendlinecollection#getcount--)|返回集合中的趋势线数量。|
||[getItem(index: number)](/javascript/api/excel/excel.charttrendlinecollection#getitem-index-)|按索引（在项目数组中的插入顺序）获取 Trendline 对象。|
||[items](/javascript/api/excel/excel.charttrendlinecollection#items)|获取此集合中已加载的子项。|
|[ChartTrendlineFormat](/javascript/api/excel/excel.charttrendlineformat)|[line](/javascript/api/excel/excel.charttrendlineformat#line)|表示图表线条格式。|
|[CustomProperty](/javascript/api/excel/excel.customproperty)|[delete()](/javascript/api/excel/excel.customproperty#delete--)|删除 custom property 对象。|
||[key](/javascript/api/excel/excel.customproperty#key)|自定义属性的键。|
||[type](/javascript/api/excel/excel.customproperty#type)|用于自定义属性的值的类型。|
||[value](/javascript/api/excel/excel.customproperty#value)|自定义属性的值。|
|[CustomPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|[add (key： string， value： any) ](/javascript/api/excel/excel.custompropertycollection#add-key--value-)|新建自定义属性或设置现有自定义属性。|
||[deleteAll () ](/javascript/api/excel/excel.custompropertycollection#deleteall--)|删除此集合中的所有自定义属性。|
||[getCount()](/javascript/api/excel/excel.custompropertycollection#getcount--)|获取自定义属性的计数。|
||[getItem(key: string)](/javascript/api/excel/excel.custompropertycollection#getitem-key-)|按键获取自定义属性对象（不区分大小写）。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.custompropertycollection#getitemornullobject-key-)|按键获取自定义属性对象（不区分大小写）。|
||[items](/javascript/api/excel/excel.custompropertycollection#items)|获取此集合中已加载的子项。|
|[DataConnectionCollection](/javascript/api/excel/excel.dataconnectioncollection)|[refreshAll () ](/javascript/api/excel/excel.dataconnectioncollection#refreshall--)|刷新集合中的所有数据连接。|
|[DocumentProperties](/javascript/api/excel/excel.documentproperties)|[author](/javascript/api/excel/excel.documentproperties#author)|工作簿的作者。|
||[category](/javascript/api/excel/excel.documentproperties#category)|工作簿的类别。|
||[comments](/javascript/api/excel/excel.documentproperties#comments)|工作簿的注释。|
||[company](/javascript/api/excel/excel.documentproperties#company)|工作簿的公司。|
||[关键字](/javascript/api/excel/excel.documentproperties#keywords)|工作簿的关键字。|
||[manager](/javascript/api/excel/excel.documentproperties#manager)|工作簿的经理。|
||[creationDate](/javascript/api/excel/excel.documentproperties#creationdate)|获取工作簿的创建日期。|
||[custom](/javascript/api/excel/excel.documentproperties#custom)|获取工作簿的自定义属性的集合。|
||[lastAuthor](/javascript/api/excel/excel.documentproperties#lastauthor)|获取工作簿的最终作者。|
||[revisionNumber](/javascript/api/excel/excel.documentproperties#revisionnumber)|获取工作簿的修订号。|
||[subject](/javascript/api/excel/excel.documentproperties#subject)|工作簿的主题。|
||[title](/javascript/api/excel/excel.documentproperties#title)|工作簿的标题。|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[formula](/javascript/api/excel/excel.nameditem#formula)|已命名项的公式。|
||[arrayValues](/javascript/api/excel/excel.nameditem#arrayvalues)|返回包含已命名项目的值和类型的对象。|
|[NamedItemArrayValues](/javascript/api/excel/excel.nameditemarrayvalues)|[types](/javascript/api/excel/excel.nameditemarrayvalues#types)|表示已命名项目数组中每个项目的类型|
||[values](/javascript/api/excel/excel.nameditemarrayvalues#values)|表示已命名项目数组中每个项目的值。|
|[Range](/javascript/api/excel/excel.range)|[getAbsoluteResizedRange (numRows： number， numColumns： number) ](/javascript/api/excel/excel.range#getabsoluteresizedrange-numrows--numcolumns-)|获取一个 Range 对象，该对象的左上单元格与当前 Range 对象相同，但具有指定的行数和列数。|
||[getImage () ](/javascript/api/excel/excel.range#getimage--)|将区域呈现为 base64 编码的 png 图像。|
||[getSurroundingRegion () ](/javascript/api/excel/excel.range#getsurroundingregion--)|返回一个 Range 对象，该对象表示此区域左上单元格的周围区域。|
||[hyperlink](/javascript/api/excel/excel.range#hyperlink)|表示当前范围的超链接。|
||[numberFormatLocal](/javascript/api/excel/excel.range#numberformatlocal)|表示基于用户的语言设置的给定范围的 Excel 号码格式代码。|
||[isEntireColumn](/javascript/api/excel/excel.range#isentirecolumn)|表示当前区域是否为整列。|
||[isEntireRow](/javascript/api/excel/excel.range#isentirerow)|表示当前区域是否为整行。|
||[showCard () ](/javascript/api/excel/excel.range#showcard--)|显示活动单元格的卡片（如果该单元格具有富值内容）。|
||[style](/javascript/api/excel/excel.range#style)|表示当前区域的样式。|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[textOrientation](/javascript/api/excel/excel.rangeformat#textorientation)|区域内所有单元格的文本方向。|
||[useStandardHeight](/javascript/api/excel/excel.rangeformat#usestandardheight)|确定对象的行高 `Range` 是否等于工作表的标准高度。|
||[useStandardWidth](/javascript/api/excel/excel.rangeformat#usestandardwidth)|指定对象的列宽是否等于工作表的标准 `Range` 宽度。|
|[RangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|[address](/javascript/api/excel/excel.rangehyperlink#address)|表示超链接的 URL 目标。|
||[documentReference](/javascript/api/excel/excel.rangehyperlink#documentreference)|表示超链接的文档引用目标。|
||[screenTip](/javascript/api/excel/excel.rangehyperlink#screentip)|表示鼠标悬停在超链接上时显示的字符串。|
||[textToDisplay](/javascript/api/excel/excel.rangehyperlink#texttodisplay)|表示区域最左上方单元格中显示的字符串。|
|[Style](/javascript/api/excel/excel.style)|[delete()](/javascript/api/excel/excel.style#delete--)|删除此样式。|
||[formulaHidden](/javascript/api/excel/excel.style#formulahidden)|指定在工作表受保护时是否隐藏公式。|
||[horizontalAlignment](/javascript/api/excel/excel.style#horizontalalignment)|表示样式水平对齐。|
||[includeAlignment](/javascript/api/excel/excel.style#includealignment)|指定样式是否包括 AutoIndent、HorizontalAlignment、VerticalAlignment、WrapText、IndentLevel 和 TextOrientation 属性。|
||[includeBorder](/javascript/api/excel/excel.style#includeborder)|指定样式是否包括 Color、ColorIndex、LineStyle 和 Weight 边框属性。|
||[includeFont](/javascript/api/excel/excel.style#includefont)|指定样式是否包括 Background、Bold、Color、ColorIndex、FontStyle、Italic、Name、Size、Strikethrough、Subscript、Superscript 和 Underline 字体属性。|
||[includeNumber](/javascript/api/excel/excel.style#includenumber)|指定样式是否包含 NumberFormat 属性。|
||[includePatterns](/javascript/api/excel/excel.style#includepatterns)|指定样式是否包括 Color、ColorIndex、InvertIfNegative、Pattern、PatternColor 和 PatternColorIndex 内部属性。|
||[includeProtection](/javascript/api/excel/excel.style#includeprotection)|指定样式是否包括 FormulaHidden 和 Locked 保护属性。|
||[indentLevel](/javascript/api/excel/excel.style#indentlevel)|0 到 250 之间的一个整数，指示样式的缩进水平。|
||[locked](/javascript/api/excel/excel.style#locked)|指定在工作表受保护时对象是否被锁定。|
||[numberFormat](/javascript/api/excel/excel.style#numberformat)|样式中数字格式的格式代码。|
||[numberFormatLocal](/javascript/api/excel/excel.style#numberformatlocal)|样式中数字格式的本地化格式代码。|
||[readingOrder](/javascript/api/excel/excel.style#readingorder)|样式中的阅读顺序。|
||[Borders](/javascript/api/excel/excel.style#borders)|四个 Border 对象的 Border 集合，表示四个边框的样式。|
||[builtIn](/javascript/api/excel/excel.style#builtin)|指定样式是否内置样式。|
||[fill](/javascript/api/excel/excel.style#fill)|样式的填充。|
||[font](/javascript/api/excel/excel.style#font)|该 Font 对象表示样式的字体。|
||[name](/javascript/api/excel/excel.style#name)|样式的名称。|
||[shrinkToFit](/javascript/api/excel/excel.style#shrinktofit)|指定文本是否自动缩小以适应可用列宽。|
||[verticalAlignment](/javascript/api/excel/excel.style#verticalalignment)|指定样式的垂直对齐方式。|
||[wrapText](/javascript/api/excel/excel.style#wraptext)|指定 Excel 是否将对象中的文本换行。|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[add(name: string)](/javascript/api/excel/excel.stylecollection#add-name-)|向集合添加新样式。|
||[getItem(name: string)](/javascript/api/excel/excel.stylecollection#getitem-name-)|按名称获取样式。|
||[items](/javascript/api/excel/excel.stylecollection#items)|获取此集合中已加载的子项。|
|[Table](/javascript/api/excel/excel.table)|[onChanged](/javascript/api/excel/excel.table#onchanged)|当单元格中的数据更改特定表时发生。|
||[onSelectionChanged](/javascript/api/excel/excel.table#onselectionchanged)|当特定表上的选定内容发生更改时发生。|
|[TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|[address](/javascript/api/excel/excel.tablechangedeventargs#address)|获取地址，该地址表示特定工作表上的表格的更改区域。|
||[changeType](/javascript/api/excel/excel.tablechangedeventargs#changetype)|获取更改类型，该类型表示 Changed 事件的触发方式。|
||[source](/javascript/api/excel/excel.tablechangedeventargs#source)|获取事件源。|
||[tableId](/javascript/api/excel/excel.tablechangedeventargs#tableid)|获取其中的数据发生更改的表格的 ID。|
||[type](/javascript/api/excel/excel.tablechangedeventargs#type)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.tablechangedeventargs#worksheetid)|获取其中的数据发生更改的工作表的 ID。|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onChanged](/javascript/api/excel/excel.tablecollection#onchanged)|在工作簿或工作表中任何表上的数据发生更改时发生。|
|[TableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|[address](/javascript/api/excel/excel.tableselectionchangedeventargs#address)|获取区域地址，该地址表示特定工作表上的表格选定区域。|
||[isInsideTable](/javascript/api/excel/excel.tableselectionchangedeventargs#isinsidetable)|指定所选内容是否位于表格内，如果 IsInsideTable 为 false，则地址将无用。|
||[tableId](/javascript/api/excel/excel.tableselectionchangedeventargs#tableid)|获取其中的选定区域发生更改的表格 ID。|
||[type](/javascript/api/excel/excel.tableselectionchangedeventargs#type)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.tableselectionchangedeventargs#worksheetid)|获取其中的选定区域发生更改的工作表的 ID。|
|[Workbook](/javascript/api/excel/excel.workbook)|[getActiveCell () ](/javascript/api/excel/excel.workbook#getactivecell--)|获取工作簿中当前处于活动状态的单元格。|
||[dataConnections](/javascript/api/excel/excel.workbook#dataconnections)|表示工作簿中所有数据连接。|
||[name](/javascript/api/excel/excel.workbook#name)|获取工作簿名称。|
||[properties](/javascript/api/excel/excel.workbook#properties)|获取工作簿属性。|
||[protection](/javascript/api/excel/excel.workbook#protection)|返回工作簿的保护对象。|
||[styles](/javascript/api/excel/excel.workbook#styles)|表示与工作簿关联的样式的集合。|
|[WorkbookProtection](/javascript/api/excel/excel.workbookprotection)|[保护 (密码？：字符串) ](/javascript/api/excel/excel.workbookprotection#protect-password-)|保护工作簿。|
||[protected](/javascript/api/excel/excel.workbookprotection#protected)|指定工作簿是否受保护。|
||[unprotect (password？： string) ](/javascript/api/excel/excel.workbookprotection#unprotect-password-)|解除保护工作簿。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[copy (positionType？： Excel.WorksheetPositionType， relativeTo？： Excel.Worksheet) ](/javascript/api/excel/excel.worksheet#copy-positiontype--relativeto-)|复制工作表，并放置于指定位置。|
||[getRangeByIndexes (startRow： number， startColumn： number， rowCount： number， columnCount： number) ](/javascript/api/excel/excel.worksheet#getrangebyindexes-startrow--startcolumn--rowcount--columncount-)|获取以特定行索引和列索引开始并跨越了一定数量的行和列的 range 对象。|
||[freezePanes](/javascript/api/excel/excel.worksheet#freezepanes)|获取一个对象，该对象可用于操作工作表上的冻结窗格。|
||[onActivated](/javascript/api/excel/excel.worksheet#onactivated)|在激活工作表时发生。|
||[onChanged](/javascript/api/excel/excel.worksheet#onchanged)|当特定工作表中的数据发生更改时发生。|
||[onDeactivated](/javascript/api/excel/excel.worksheet#ondeactivated)|在工作表停用时发生。|
||[onSelectionChanged](/javascript/api/excel/excel.worksheet#onselectionchanged)|当特定工作表上的选定内容发生更改时发生。|
||[standardHeight](/javascript/api/excel/excel.worksheet#standardheight)|返回工作表中所有行的标准（默认）行高，以磅为单位。|
||[standardWidth](/javascript/api/excel/excel.worksheet#standardwidth)|指定工作表中 (列) 列的默认列宽。|
||[tabColor](/javascript/api/excel/excel.worksheet#tabcolor)|工作表的选项卡颜色。|
|[WorksheetActivatedEventArgs](/javascript/api/excel/excel.worksheetactivatedeventargs)|[type](/javascript/api/excel/excel.worksheetactivatedeventargs#type)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.worksheetactivatedeventargs#worksheetid)|获取已启用的工作表的 ID。|
|[WorksheetAddedEventArgs](/javascript/api/excel/excel.worksheetaddedeventargs)|[source](/javascript/api/excel/excel.worksheetaddedeventargs#source)|获取事件源。|
||[type](/javascript/api/excel/excel.worksheetaddedeventargs#type)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.worksheetaddedeventargs#worksheetid)|获取已添加至工作簿的工作表的 ID。|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[address](/javascript/api/excel/excel.worksheetchangedeventargs#address)|获取区域地址，该地址表示特定工作表上的更改区域。|
||[changeType](/javascript/api/excel/excel.worksheetchangedeventargs#changetype)|获取更改类型，该类型表示 Changed 事件的触发方式。|
||[source](/javascript/api/excel/excel.worksheetchangedeventargs#source)|获取事件源。|
||[type](/javascript/api/excel/excel.worksheetchangedeventargs#type)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.worksheetchangedeventargs#worksheetid)|获取其中的数据发生更改的工作表的 ID。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onActivated](/javascript/api/excel/excel.worksheetcollection#onactivated)|在工作簿中任何工作表被激活时发生。|
||[onAdded](/javascript/api/excel/excel.worksheetcollection#onadded)|在将新工作表添加到工作簿时发生。|
||[onDeactivated](/javascript/api/excel/excel.worksheetcollection#ondeactivated)|在工作簿中任何工作表被停用时发生。|
||[onDeleted](/javascript/api/excel/excel.worksheetcollection#ondeleted)|从工作簿中删除工作表时发生。|
|[WorksheetDeactivatedEventArgs](/javascript/api/excel/excel.worksheetdeactivatedeventargs)|[type](/javascript/api/excel/excel.worksheetdeactivatedeventargs#type)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.worksheetdeactivatedeventargs#worksheetid)|获取已停用的工作表的 ID。|
|[WorksheetDeletedEventArgs](/javascript/api/excel/excel.worksheetdeletedeventargs)|[source](/javascript/api/excel/excel.worksheetdeletedeventargs#source)|获取事件源。|
||[type](/javascript/api/excel/excel.worksheetdeletedeventargs#type)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.worksheetdeletedeventargs#worksheetid)|获取已从工作簿删除的工作表的 ID。|
|[WorksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|[freezeAt (frozenRange： Range \| string) ](/javascript/api/excel/excel.worksheetfreezepanes#freezeat-frozenrange-)|设置活动工作表视图中的冻结单元格。|
||[freezeColumns (计数？： number) ](/javascript/api/excel/excel.worksheetfreezepanes#freezecolumns-count-)|就地冻结工作表的第一列。|
||[freezeRows (计数？： number) ](/javascript/api/excel/excel.worksheetfreezepanes#freezerows-count-)|就地冻结工作表的首行。|
||[getLocation()](/javascript/api/excel/excel.worksheetfreezepanes#getlocation--)|获取用于描述活动工作表视图中的冻结单元格的区域。|
||[getLocationOrNullObject () ](/javascript/api/excel/excel.worksheetfreezepanes#getlocationornullobject--)|获取用于描述活动工作表视图中的冻结单元格的区域。|
||[unfreeze () ](/javascript/api/excel/excel.worksheetfreezepanes#unfreeze--)|移除工作表中的所有冻结窗格。|
|[WorksheetProtection](/javascript/api/excel/excel.worksheetprotection)|[unprotect (password？： string) ](/javascript/api/excel/excel.worksheetprotection#unprotect-password-)|解除对 worksheet 的保护。|
|[WorksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|[allowEditObjects](/javascript/api/excel/excel.worksheetprotectionoptions#alloweditobjects)|表示允许编辑对象的工作表保护选项。|
||[allowEditScenarios](/javascript/api/excel/excel.worksheetprotectionoptions#alloweditscenarios)|表示允许编辑应用场景的工作表保护选项。|
||[selectionMode](/javascript/api/excel/excel.worksheetprotectionoptions#selectionmode)|表示选择模式的工作表保护选项。|
|[WorksheetSelectionChangedEventArgs](/javascript/api/excel/excel.worksheetselectionchangedeventargs)|[address](/javascript/api/excel/excel.worksheetselectionchangedeventargs#address)|获取区域地址，该地址表示特定工作表上的选定区域。|
||[type](/javascript/api/excel/excel.worksheetselectionchangedeventargs#type)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.worksheetselectionchangedeventargs#worksheetid)|获取其中的选定区域发生更改的工作表的 ID。|
