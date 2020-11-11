| Class | 域 | 说明 |
|:---|:---|:---|
|[BindingCollection](/javascript/api/excel/excel.bindingcollection)|[getCount()](/javascript/api/excel/excel.bindingcollection#getcount--)|获取集合中的绑定数量。|
||[getItemOrNullObject(id: string)](/javascript/api/excel/excel.bindingcollection#getitemornullobject-id-)|按 ID 获取绑定对象。|
|[ChartCollection](/javascript/api/excel/excel.chartcollection)|[getCount()](/javascript/api/excel/excel.chartcollection#getcount--)|返回工作表中的图表数。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.chartcollection#getitemornullobject-name-)|使用图表名称获取图表。|
|[ChartPointsCollection](/javascript/api/excel/excel.chartpointscollection)|[getCount()](/javascript/api/excel/excel.chartpointscollection#getcount--)|返回系列中的图表点数。|
|[ChartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|[getCount()](/javascript/api/excel/excel.chartseriescollection#getcount--)|返回集合中的系列数量。|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[comment](/javascript/api/excel/excel.nameditem#comment)|指定与此名称关联的注释。|
||[delete()](/javascript/api/excel/excel.nameditem#delete--)|删除给定的名称。|
||[getRangeOrNullObject()](/javascript/api/excel/excel.nameditem#getrangeornullobject--)|返回与名称相关的 range 对象。|
||[scope](/javascript/api/excel/excel.nameditem#scope)|指定名称的范围是工作簿还是特定工作表。|
||[worksheet](/javascript/api/excel/excel.nameditem#worksheet)|返回已命名项限定到的工作表。|
||[worksheetOrNullObject](/javascript/api/excel/excel.nameditem#worksheetornullobject)|返回已命名项限定到的工作表。|
|[NamedItemCollection](/javascript/api/excel/excel.nameditemcollection)|[添加 (名称： string，reference： Range \| string，comment？： string) ](/javascript/api/excel/excel.nameditemcollection#add-name--reference--comment-)|将新名称添加到给定范围的集合。|
||[addFormulaLocal (名称： string，formula： string，comment？： string) ](/javascript/api/excel/excel.nameditemcollection#addformulalocal-name--formula--comment-)|使用用户的公式区域设置，将新名称添加到给定范围的集合。|
||[getCount()](/javascript/api/excel/excel.nameditemcollection#getcount--)|获取集合中已命名项的数量。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.nameditemcollection#getitemornullobject-name-)|使用其名称获取 NamedItem 对象。|
|[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|[getCount()](/javascript/api/excel/excel.pivottablecollection#getcount--)|获取集合中的数据透视表的数量。|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivottablecollection#getitemornullobject-name-)|按名称获取 PivotTable 对象。|
|[Range](/javascript/api/excel/excel.range)|[getIntersectionOrNullObject (anotherRange： Range \| string) ](/javascript/api/excel/excel.range#getintersectionornullobject-anotherrange-)|获取表示指定区域的矩形交集的 range 对象。|
||[getUsedRangeOrNullObject (valuesOnly？： boolean) ](/javascript/api/excel/excel.range#getusedrangeornullobject-valuesonly-)|返回指定 range 对象的所用区域。|
|[RangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|[getCount()](/javascript/api/excel/excel.rangeviewcollection#getcount--)|获取集合中 RangeView 对象的数量。|
|[设置](/javascript/api/excel/excel.setting)|[delete()](/javascript/api/excel/excel.setting#delete--)|删除 Setting 对象。|
||[key](/javascript/api/excel/excel.setting#key)|表示设置的 id 的键。|
||[value](/javascript/api/excel/excel.setting#value)|表示为此设置存储的值。|
|[SettingCollection](/javascript/api/excel/excel.settingcollection)|[add (key： string，value： string \| number \| boolean \| Date \| Array <any> \| any) ](/javascript/api/excel/excel.settingcollection#add-key--value-)|设置指定的 Setting 对象，或将其添加到工作簿中。|
||[getCount()](/javascript/api/excel/excel.settingcollection#getcount--)|获取集合中的 Setting 对象的数量。|
||[getItem(key: string)](/javascript/api/excel/excel.settingcollection#getitem-key-)|按键获取 Setting 项。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.settingcollection#getitemornullobject-key-)|按键获取 Setting 项。|
||[items](/javascript/api/excel/excel.settingcollection#items)|获取此集合中已加载的子项。|
||[onSettingsChanged](/javascript/api/excel/excel.settingcollection#onsettingschanged)|当文档中的设置变化时发生。|
|[SettingsChangedEventArgs](/javascript/api/excel/excel.settingschangedeventargs)|[设置](/javascript/api/excel/excel.settingschangedeventargs#settings)|获取表示引发了 SettingsChanged 事件的 binding 的 setting 对象。|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[getCount()](/javascript/api/excel/excel.tablecollection#getcount--)|获取集合中的表数量。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.tablecollection#getitemornullobject-key-)|按名称或 ID 获取表。|
|[TableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|[getCount()](/javascript/api/excel/excel.tablecolumncollection#getcount--)|获取表中的列数。|
||[getItemOrNullObject (项：数字 \| 字符串) ](/javascript/api/excel/excel.tablecolumncollection#getitemornullobject-key-)|按名称或 ID 获取 column 对象。|
|[TableRowCollection](/javascript/api/excel/excel.tablerowcollection)|[getCount()](/javascript/api/excel/excel.tablerowcollection#getcount--)|获取表格中的行数。|
|[Workbook](/javascript/api/excel/excel.workbook)|[设置](/javascript/api/excel/excel.workbook#settings)|表示一组与 workbook 相关联的 setting 对象。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[getUsedRangeOrNullObject (valuesOnly？： boolean) ](/javascript/api/excel/excel.worksheet#getusedrangeornullobject-valuesonly-)|使用的区域是包含分配了值或格式化的任何单元格的最小区域。|
||[名称](/javascript/api/excel/excel.worksheet#names)|一组范围限定到当前工作表的名称。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[getCount (visibleOnly？： boolean) ](/javascript/api/excel/excel.worksheetcollection#getcount-visibleonly-)|获取集合中的工作表数量。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.worksheetcollection#getitemornullobject-key-)|按 Worksheet 对象的名称或 ID 获取此对象。|
