| Class | 域 | 说明 |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[suspendApiCalculationUntilNextSync ( # B1 ](/javascript/api/excel/excel.application#suspendapicalculationuntilnextsync--)|在下一次调用“context.sync()”前暂停计算。|
|[CellValueConditionalFormat](/javascript/api/excel/excel.cellvalueconditionalformat)|[format](/javascript/api/excel/excel.cellvalueconditionalformat#format)|返回一个 format 对象，该对象封装条件格式字体、填充、边框和其他属性。|
||[标尺](/javascript/api/excel/excel.cellvalueconditionalformat#rule)|指定此条件格式的 Rule 对象。|
|[ColorScaleConditionalFormat](/javascript/api/excel/excel.colorscaleconditionalformat)|[criteria](/javascript/api/excel/excel.colorscaleconditionalformat#criteria)|色阶的条件。|
||[threeColorScale](/javascript/api/excel/excel.colorscaleconditionalformat#threecolorscale)|如果为 true，则色阶将具有三个点 (最小、中点、最大) ，否则它将有两个 (最小值，最大) 。|
|[ConditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule)|[formula1](/javascript/api/excel/excel.conditionalcellvaluerule#formula1)|如果需要，公式可对条件格式规则进行求值。|
||[formula2](/javascript/api/excel/excel.conditionalcellvaluerule#formula2)|如果需要，公式可对条件格式规则进行求值。|
||[operator](/javascript/api/excel/excel.conditionalcellvaluerule#operator)|单元格数值条件格式的运算符。|
|[ConditionalColorScaleCriteria](/javascript/api/excel/excel.conditionalcolorscalecriteria)|[maximum](/javascript/api/excel/excel.conditionalcolorscalecriteria#maximum)|最大点色阶条件。|
||[放置](/javascript/api/excel/excel.conditionalcolorscalecriteria#midpoint)|色阶为 3 色阶时的中点色阶条件。|
||[minimum](/javascript/api/excel/excel.conditionalcolorscalecriteria#minimum)|最小点色阶条件。|
|[ConditionalColorScaleCriterion](/javascript/api/excel/excel.conditionalcolorscalecriterion)|[color](/javascript/api/excel/excel.conditionalcolorscalecriterion#color)|色阶颜色的 HTML 颜色代码表示形式 (例如，#FF0000 代表红色) 。|
||[formula](/javascript/api/excel/excel.conditionalcolorscalecriterion#formula)|数字、公式或 null（如果类型为 LowestValue）。|
||[type](/javascript/api/excel/excel.conditionalcolorscalecriterion#type)|条件条件公式应基于什么。|
|[ConditionalDataBarNegativeFormat](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|[borderColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#bordercolor)|表示窗体 #RRGGBB（例如 "FFA500"）的边框线条颜色或作为已命名的 HTML 颜色（例如 "orange"）的 HTML 颜色代码。|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#fillcolor)|代表窗体 #RRGGBB (的填充颜色的 HTML 颜色代码，例如，"FFA500" ) 或作为命名的 HTML 颜色 (例如 "橙色" ) 。|
||[matchPositiveBorderColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#matchpositivebordercolor)|指定负 DataBar 的边框颜色是否与正 DataBar 相同。|
||[matchPositiveFillColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#matchpositivefillcolor)|指定负 DataBar 的填充颜色是否与正 DataBar 相同。|
|[ConditionalDataBarPositiveFormat](/javascript/api/excel/excel.conditionaldatabarpositiveformat)|[borderColor](/javascript/api/excel/excel.conditionaldatabarpositiveformat#bordercolor)|表示窗体 #RRGGBB（例如 "FFA500"）的边框线条颜色或作为已命名的 HTML 颜色（例如 "orange"）的 HTML 颜色代码。|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarpositiveformat#fillcolor)|代表窗体 #RRGGBB (的填充颜色的 HTML 颜色代码，例如，"FFA500" ) 或作为命名的 HTML 颜色 (例如 "橙色" ) 。|
||[gradientFill](/javascript/api/excel/excel.conditionaldatabarpositiveformat#gradientfill)|指定 DataBar 是否具有渐变。|
|[ConditionalDataBarRule](/javascript/api/excel/excel.conditionaldatabarrule)|[formula](/javascript/api/excel/excel.conditionaldatabarrule#formula)|如果需要，公式可对 databar 规则进行求值。|
||[type](/javascript/api/excel/excel.conditionaldatabarrule#type)|Databar 的规则类型。|
|[ConditionalFormat](/javascript/api/excel/excel.conditionalformat)|[delete()](/javascript/api/excel/excel.conditionalformat#delete--)|删除此条件格式。|
||[getRange()](/javascript/api/excel/excel.conditionalformat#getrange--)|返回应用条件格式的范围。|
||[getRangeOrNullObject()](/javascript/api/excel/excel.conditionalformat#getrangeornullobject--)|返回条件格式应用于的区域; 或者，如果将条件格式应用于多个区域，则返回 null 对象。|
||[priority](/javascript/api/excel/excel.conditionalformat#priority)|条件格式集合中当前存在此条件格式的优先级 (或索引) 。|
||[cellValue](/javascript/api/excel/excel.conditionalformat#cellvalue)|如果当前条件格式为 CellValue 类型，则返回单元格值条件格式属性。|
||[cellValueOrNullObject](/javascript/api/excel/excel.conditionalformat#cellvalueornullobject)|如果当前条件格式为 CellValue 类型，则返回单元格值条件格式属性。|
||[色阶](/javascript/api/excel/excel.conditionalformat#colorscale)|如果当前条件格式为色阶类型，则返回色阶条件格式属性。|
||[colorScaleOrNullObject](/javascript/api/excel/excel.conditionalformat#colorscaleornullobject)|如果当前条件格式为色阶类型，则返回色阶条件格式属性。|
||[自](/javascript/api/excel/excel.conditionalformat#custom)|如果当前条件格式为自定义类型，则返回自定义条件格式属性。|
||[customOrNullObject](/javascript/api/excel/excel.conditionalformat#customornullobject)|如果当前条件格式为自定义类型，则返回自定义条件格式属性。|
||[dataBar](/javascript/api/excel/excel.conditionalformat#databar)|如果当前条件格式为数据栏，则返回数据条属性。|
||[dataBarOrNullObject](/javascript/api/excel/excel.conditionalformat#databarornullobject)|如果当前条件格式为数据栏，则返回数据条属性。|
||[iconSet](/javascript/api/excel/excel.conditionalformat#iconset)|如果当前条件格式为 IconSet 类型，则返回 IconSet 条件格式属性。|
||[iconSetOrNullObject](/javascript/api/excel/excel.conditionalformat#iconsetornullobject)|如果当前条件格式为 IconSet 类型，则返回 IconSet 条件格式属性。|
||[id](/javascript/api/excel/excel.conditionalformat#id)|当前 ConditionalFormatCollection 内的条件格式的优先级。|
||[好](/javascript/api/excel/excel.conditionalformat#preset)|返回预设条件的条件格式。|
||[presetOrNullObject](/javascript/api/excel/excel.conditionalformat#presetornullobject)|返回预设条件的条件格式。|
||[textComparison](/javascript/api/excel/excel.conditionalformat#textcomparison)|如果当前条件格式是文本类型，则返回特定的文本条件格式属性。|
||[textComparisonOrNullObject](/javascript/api/excel/excel.conditionalformat#textcomparisonornullobject)|如果当前条件格式是文本类型，则返回特定的文本条件格式属性。|
||[topBottom](/javascript/api/excel/excel.conditionalformat#topbottom)|如果当前条件格式为 TopBottom 类型，则返回 Top/底端条件格式属性。|
||[topBottomOrNullObject](/javascript/api/excel/excel.conditionalformat#topbottomornullobject)|如果当前条件格式为 TopBottom 类型，则返回 Top/底端条件格式属性。|
||[type](/javascript/api/excel/excel.conditionalformat#type)|一种条件格式。|
||[stopIfTrue](/javascript/api/excel/excel.conditionalformat#stopiftrue)|如果满足此条件格式的条件，则不会有任何低优先级格式应在此单元格上生效。|
|[ConditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|[添加 (类型： ConditionalFormatType) ](/javascript/api/excel/excel.conditionalformatcollection#add-type-)|将新的条件格式添加到集合中的第一个/最高优先级处。|
||[clearAll ( # B1 ](/javascript/api/excel/excel.conditionalformatcollection#clearall--)|清除当前指定区域中处于活动状态的所有条件格式。|
||[getCount()](/javascript/api/excel/excel.conditionalformatcollection#getcount--)|返回工作簿中的条件格式数。|
||[getItem(id: string)](/javascript/api/excel/excel.conditionalformatcollection#getitem-id-)|返回给定 ID 的条件格式。|
||[getItemAt(index: number)](/javascript/api/excel/excel.conditionalformatcollection#getitemat-index-)|返回给定索引处的条件格式。|
||[items](/javascript/api/excel/excel.conditionalformatcollection#items)|获取此集合中已加载的子项。|
|[ConditionalFormatRule](/javascript/api/excel/excel.conditionalformatrule)|[formula](/javascript/api/excel/excel.conditionalformatrule#formula)|如果需要，公式可对条件格式规则进行求值。|
||[formulaLocal](/javascript/api/excel/excel.conditionalformatrule#formulalocal)|如果需要，公式可采用用户的语言对条件格式规则进行求值。|
||[formulaR1C1](/javascript/api/excel/excel.conditionalformatrule#formular1c1)|如果需要，公式可采用 R1C1 表示法对条件格式规则进行求值。|
|[ConditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|[customIcon](/javascript/api/excel/excel.conditionaliconcriterion#customicon)|如果与默认 IconSet 不同，返回当前条件的自定义图标，否则将返回 null。|
||[formula](/javascript/api/excel/excel.conditionaliconcriterion#formula)|取决于类型的数字或公式。|
||[operator](/javascript/api/excel/excel.conditionaliconcriterion#operator)|图标条件格式的每个规则类型的 GreaterThan 或 GreaterThanOrEqual。|
||[type](/javascript/api/excel/excel.conditionaliconcriterion#type)|应基于的图标条件公式。|
|[ConditionalPresetCriteriaRule](/javascript/api/excel/excel.conditionalpresetcriteriarule)|[依据](/javascript/api/excel/excel.conditionalpresetcriteriarule#criterion)|条件格式的条件。|
|[ConditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|[color](/javascript/api/excel/excel.conditionalrangeborder#color)|表示窗体 #RRGGBB（例如 "FFA500"）的边框线条颜色或作为已命名的 HTML 颜色（例如 "orange"）的 HTML 颜色代码。|
||[sideIndex](/javascript/api/excel/excel.conditionalrangeborder#sideindex)|指示边框的特定边的常量值。|
||[style](/javascript/api/excel/excel.conditionalrangeborder#style)|线条样式的常量之一，指定边框的线条样式。|
|[ConditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|[getItem (索引： ConditionalRangeBorderIndex) ](/javascript/api/excel/excel.conditionalrangebordercollection#getitem-index-)|使用其名称获取 border 对象|
||[getItemAt(index: number)](/javascript/api/excel/excel.conditionalrangebordercollection#getitemat-index-)|使用其索引获取 border 对象|
||[bottom](/javascript/api/excel/excel.conditionalrangebordercollection#bottom)|获取下边框。|
||[count](/javascript/api/excel/excel.conditionalrangebordercollection#count)|集合中的 border 对象数量。|
||[items](/javascript/api/excel/excel.conditionalrangebordercollection#items)|获取此集合中已加载的子项。|
||[left](/javascript/api/excel/excel.conditionalrangebordercollection#left)|获取左边框。|
||[right](/javascript/api/excel/excel.conditionalrangebordercollection#right)|获取右边框。|
||[top](/javascript/api/excel/excel.conditionalrangebordercollection#top)|获取上边框。|
|[ConditionalRangeFill](/javascript/api/excel/excel.conditionalrangefill)|[clear()](/javascript/api/excel/excel.conditionalrangefill#clear--)|重置填充。|
||[color](/javascript/api/excel/excel.conditionalrangefill#color)|代表填充颜色的 HTML 颜色代码 #RRGGBB (例如，"FFA500" ) 或作为命名的 HTML 颜色 (例如 "橙色" ) 。|
|[ConditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|[bold](/javascript/api/excel/excel.conditionalrangefont#bold)|指定字体是否为粗体。|
||[clear()](/javascript/api/excel/excel.conditionalrangefont#clear--)|重置字体格式。|
||[color](/javascript/api/excel/excel.conditionalrangefont#color)|文本颜色的 HTML 颜色代码表示 (例如，#FF0000 代表红色) 。|
||[italic](/javascript/api/excel/excel.conditionalrangefont#italic)|指定字体是否为斜体。|
||[strikethrough](/javascript/api/excel/excel.conditionalrangefont#strikethrough)|指定字体的删除线状态。|
||[underline](/javascript/api/excel/excel.conditionalrangefont#underline)|应用于字体的下划线类型。|
|[ConditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|[numberFormat](/javascript/api/excel/excel.conditionalrangeformat#numberformat)|表示给定范围的 Excel 数字格式代码。|
||[Borders](/javascript/api/excel/excel.conditionalrangeformat#borders)|应用于整体条件格式范围的 border 对象的集合。|
||[fill](/javascript/api/excel/excel.conditionalrangeformat#fill)|返回在整体条件格式范围上定义的 fill 对象。|
||[font](/javascript/api/excel/excel.conditionalrangeformat#font)|返回在整体条件格式区域上定义的 font 对象。|
|[ConditionalTextComparisonRule](/javascript/api/excel/excel.conditionaltextcomparisonrule)|[operator](/javascript/api/excel/excel.conditionaltextcomparisonrule#operator)|文本条件格式的运算符。|
||[text](/javascript/api/excel/excel.conditionaltextcomparisonrule#text)|条件格式的文本值。|
|[ConditionalTopBottomRule](/javascript/api/excel/excel.conditionaltopbottomrule)|[rank](/javascript/api/excel/excel.conditionaltopbottomrule#rank)|1 和 1000 之间的数字排名或 1 和 100 之间的百分比排名。|
||[type](/javascript/api/excel/excel.conditionaltopbottomrule#type)|根据顶部或底部排名设置值的格式。|
|[CustomConditionalFormat](/javascript/api/excel/excel.customconditionalformat)|[format](/javascript/api/excel/excel.customconditionalformat#format)|返回一个 format 对象，该对象封装条件格式字体、填充、边框和其他属性。|
||[标尺](/javascript/api/excel/excel.customconditionalformat#rule)|指定此条件格式的 Rule 对象。|
|[DataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|[axisColor](/javascript/api/excel/excel.databarconditionalformat#axiscolor)|代表坐标轴颜色的 HTML 颜色代码，格式 #RRGGBB (，例如，"FFA500" ) 或作为命名的 HTML 颜色 (例如，"橙色" ) 。|
||[axisFormat](/javascript/api/excel/excel.databarconditionalformat#axisformat)|为 Excel 数据栏确定轴的方式的表示形式。|
||[barDirection](/javascript/api/excel/excel.databarconditionalformat#bardirection)|指定数据条图形的基于方向。|
||[lowerBoundRule](/javascript/api/excel/excel.databarconditionalformat#lowerboundrule)|构成数据栏的下限（以及如何计算，如果适用）的规则。|
||[negativeFormat](/javascript/api/excel/excel.databarconditionalformat#negativeformat)|在 Excel 数据栏中的轴左侧的所有值的表示形式。|
||[positiveFormat](/javascript/api/excel/excel.databarconditionalformat#positiveformat)|在 Excel 数据栏中的轴右侧的所有值的表示形式。|
||[showDataBarOnly](/javascript/api/excel/excel.databarconditionalformat#showdatabaronly)|如果为 true，则对应用数据栏的单元格隐藏值。|
||[upperBoundRule](/javascript/api/excel/excel.databarconditionalformat#upperboundrule)|构成数据栏的上限（以及如何计算，如果适用）的规则。|
|[IconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|[criteria](/javascript/api/excel/excel.iconsetconditionalformat#criteria)|用于条件图标的规则和潜在自定义图标的条件和 IconSets 的数组。|
||[reverseIconOrder](/javascript/api/excel/excel.iconsetconditionalformat#reverseiconorder)|如果为 true，则反转 IconSet 的图标订单。|
||[showIconOnly](/javascript/api/excel/excel.iconsetconditionalformat#showicononly)|如果为 true，则隐藏值并仅显示图标。|
||[style](/javascript/api/excel/excel.iconsetconditionalformat#style)|如果设置，则显示条件格式的 IconSet 选项。|
|[PresetCriteriaConditionalFormat](/javascript/api/excel/excel.presetcriteriaconditionalformat)|[format](/javascript/api/excel/excel.presetcriteriaconditionalformat#format)|返回一个 format 对象，该对象封装条件格式字体、填充、边框和其他属性。|
||[标尺](/javascript/api/excel/excel.presetcriteriaconditionalformat#rule)|条件格式的规则。|
|[Range](/javascript/api/excel/excel.range)|[calculate()](/javascript/api/excel/excel.range#calculate--)|计算工作表上的单元格区域。|
||[conditionalFormats](/javascript/api/excel/excel.range#conditionalformats)|与该范围相交的 ConditionalFormats 的集合。|
|[TextConditionalFormat](/javascript/api/excel/excel.textconditionalformat)|[format](/javascript/api/excel/excel.textconditionalformat#format)|返回一个 format 对象，该对象封装条件格式的字体、填充、边框和其他属性。|
||[标尺](/javascript/api/excel/excel.textconditionalformat#rule)|条件格式的规则。|
|[TopBottomConditionalFormat](/javascript/api/excel/excel.topbottomconditionalformat)|[format](/javascript/api/excel/excel.topbottomconditionalformat#format)|返回一个 format 对象，该对象封装条件格式的字体、填充、边框和其他属性。|
||[标尺](/javascript/api/excel/excel.topbottomconditionalformat#rule)|顶部/底部条件格式的条件。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[计算 (markAllDirty： boolean) ](/javascript/api/excel/excel.worksheet#calculate-markalldirty-)|计算工作表上的所有单元格。|
