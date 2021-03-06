| 类 | 域 | 说明 |
|:---|:---|:---|
|[AutoFilter](/javascript/api/excel/excel.autofilter)|[clearColumnCriteria (columnIndex： number) ](/javascript/api/excel/excel.autofilter#clearcolumncriteria-columnindex-)|清除自动筛选器的筛选条件。|
|[Comment](/javascript/api/excel/excel.comment)|[assignTask (：Identity) ](/javascript/api/excel/excel.comment#assigntask-assignee-)|将附加到注释的任务作为收件人分配给给定用户。|
||[getTask () ](/javascript/api/excel/excel.comment#gettask--)|获取与此注释关联的任务。|
||[getTaskOrNullObject () ](/javascript/api/excel/excel.comment#gettaskornullobject--)|获取与此注释关联的任务。|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[getItemOrNullObject (commentId： string) ](/javascript/api/excel/excel.commentcollection#getitemornullobject-commentid-)|根据其 ID 从集合中获取批注。|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[assignTask (：Identity) ](/javascript/api/excel/excel.commentreply#assigntask-assignee-)|将附加到注释的任务作为唯一的代理人分配给给定用户。|
||[getTask () ](/javascript/api/excel/excel.commentreply#gettask--)|获取与此注释回复线程关联的任务。|
||[getTaskOrNullObject () ](/javascript/api/excel/excel.commentreply#gettaskornullobject--)|获取与此注释回复线程关联的任务。|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[getItemOrNullObject (commentReplyId： string) ](/javascript/api/excel/excel.commentreplycollection#getitemornullobject-commentreplyid-)|返回由其 ID 标识的批注回复。|
|[ConditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|[getItemOrNullObject(id: string)](/javascript/api/excel/excel.conditionalformatcollection#getitemornullobject-id-)|返回由其 ID 标识的条件格式。|
|[DocumentTask](/javascript/api/excel/excel.documenttask)|[percentComplete](/javascript/api/excel/excel.documenttask#percentcomplete)|指定任务的完成百分比。|
||[priority](/javascript/api/excel/excel.documenttask#priority)|指定任务的优先级。|
||[assignees](/javascript/api/excel/excel.documenttask#assignees)|返回任务委派者的集合。|
||[更改](/javascript/api/excel/excel.documenttask#changes)|获取任务的更改记录。|
||[comment](/javascript/api/excel/excel.documenttask#comment)|获取与任务关联的注释。|
||[completedBy](/javascript/api/excel/excel.documenttask#completedby)|获取已完成该任务的最新用户。|
||[completedDateTime](/javascript/api/excel/excel.documenttask#completeddatetime)|获取任务的完成日期和时间。|
||[createdBy](/javascript/api/excel/excel.documenttask#createdby)|获取创建任务的用户。|
||[createdDateTime](/javascript/api/excel/excel.documenttask#createddatetime)|获取任务的创建日期和时间。|
||[id](/javascript/api/excel/excel.documenttask#id)|获取任务的 ID。|
||[setStartAndDueDateTime (startDateTime： Date， dueDateTime： Date) ](/javascript/api/excel/excel.documenttask#setstartandduedatetime-startdatetime--duedatetime-)|更改任务的开始日期和截止日期。|
||[startAndDueDateTime](/javascript/api/excel/excel.documenttask#startandduedatetime)|获取或设置任务应开始和到期的日期和时间。|
||[title](/javascript/api/excel/excel.documenttask#title)|指定任务的标题。|
|[DocumentTaskChange](/javascript/api/excel/excel.documenttaskchange)|[被分配者](/javascript/api/excel/excel.documenttaskchange#assignee)|表示分配给更改记录类型的任务的用户，或者从更改记录类型的任务中取消 `assign` `unassign` 分配的用户。|
||[changedBy](/javascript/api/excel/excel.documenttaskchange#changedby)|表示创建或更改任务的用户。|
||[commentId](/javascript/api/excel/excel.documenttaskchange#commentid)|表示任务 `Comment` 更改定位到 `CommentReply` 的或其 ID。|
||[createdDateTime](/javascript/api/excel/excel.documenttaskchange#createddatetime)|表示任务更改记录的创建日期和时间。|
||[dueDateTime](/javascript/api/excel/excel.documenttaskchange#duedatetime)|表示任务的截止日期和时间，以 UTC 时区表示。|
||[id](/javascript/api/excel/excel.documenttaskchange#id)|任务更改记录的 ID。|
||[percentComplete](/javascript/api/excel/excel.documenttaskchange#percentcomplete)|表示任务的完成百分比。|
||[priority](/javascript/api/excel/excel.documenttaskchange#priority)|表示任务的优先级。|
||[startDateTime](/javascript/api/excel/excel.documenttaskchange#startdatetime)|表示任务的开始日期和时间，以 UTC 时区表示。|
||[title](/javascript/api/excel/excel.documenttaskchange#title)|表示任务的标题。|
||[type](/javascript/api/excel/excel.documenttaskchange#type)|表示任务更改记录的操作类型。|
||[undoHistoryId](/javascript/api/excel/excel.documenttaskchange#undohistoryid)|表示 `DocumentTaskChange.id` 对更改记录类型撤消 `undo` 的属性。|
|[DocumentTaskChangeCollection](/javascript/api/excel/excel.documenttaskchangecollection)|[getCount()](/javascript/api/excel/excel.documenttaskchangecollection#getcount--)|获取任务集合中的更改记录数。|
||[getItemAt(index: number)](/javascript/api/excel/excel.documenttaskchangecollection#getitemat-index-)|使用任务更改记录在集合中的索引获取该记录。|
||[items](/javascript/api/excel/excel.documenttaskchangecollection#items)|获取此集合中已加载的子项。|
|[DocumentTaskCollection](/javascript/api/excel/excel.documenttaskcollection)|[getCount()](/javascript/api/excel/excel.documenttaskcollection#getcount--)|获取集合中的任务数。|
||[getItem(key: string)](/javascript/api/excel/excel.documenttaskcollection#getitem-key-)|使用其 ID 获取任务。|
||[getItemAt(index: number)](/javascript/api/excel/excel.documenttaskcollection#getitemat-index-)|按任务在集合中的索引获取任务。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.documenttaskcollection#getitemornullobject-key-)|使用其 ID 获取任务。|
||[items](/javascript/api/excel/excel.documenttaskcollection#items)|获取此集合中已加载的子项。|
|[DocumentTaskSchedule](/javascript/api/excel/excel.documenttaskschedule)|[dueDateTime](/javascript/api/excel/excel.documenttaskschedule#duedatetime)|获取任务到期的日期和时间。|
||[startDateTime](/javascript/api/excel/excel.documenttaskschedule#startdatetime)|获取任务应开始的日期和时间。|
|[FormulaChangedEventDetail](/javascript/api/excel/excel.formulachangedeventdetail)|[cellAddress](/javascript/api/excel/excel.formulachangedeventdetail#celladdress)|包含已更改公式的单元格的地址。|
||[previousFormula](/javascript/api/excel/excel.formulachangedeventdetail#previousformula)|表示更改前一个公式。|
|[GroupShapeCollection](/javascript/api/excel/excel.groupshapecollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.groupshapecollection#getitemornullobject-key-)|使用形状的名称或 ID 获取形状。|
|[标识](/javascript/api/excel/excel.identity)|[displayName](/javascript/api/excel/excel.identity#displayname)|表示用户的显示名称。|
||[email](/javascript/api/excel/excel.identity#email)|表示用户的电子邮件地址。|
||[id](/javascript/api/excel/excel.identity#id)|表示用户的唯一 ID。|
|[IdentityCollection](/javascript/api/excel/excel.identitycollection)|[添加 (：标识) ](/javascript/api/excel/excel.identitycollection#add-assignee-)|向集合添加用户标识。|
||[clear()](/javascript/api/excel/excel.identitycollection#clear--)|从集合中删除所有用户标识。|
||[getCount()](/javascript/api/excel/excel.identitycollection#getcount--)|获取集合中项的数目。|
||[getItemAt(index: number)](/javascript/api/excel/excel.identitycollection#getitemat-index-)|使用文档在集合中的索引获取文档用户标识。|
||[删除 (：标识) ](/javascript/api/excel/excel.identitycollection#remove-assignee-)|从集合中删除用户标识。|
|[InsertWorksheetOptions](/javascript/api/excel/excel.insertworksheetoptions)|[positionType](/javascript/api/excel/excel.insertworksheetoptions#positiontype)|新工作表的当前工作簿中的插入位置。|
||[relativeTo](/javascript/api/excel/excel.insertworksheetoptions#relativeto)|引用参数的当前工作簿中的 `WorksheetPositionType` 工作表。|
||[sheetNamesToInsert](/javascript/api/excel/excel.insertworksheetoptions#sheetnamestoinsert)|要插入的单个工作表的名称。|
|[LinkedDataType](/javascript/api/excel/excel.linkeddatatype)|[dataProvider](/javascript/api/excel/excel.linkeddatatype#dataprovider)|链接数据提供程序的数据提供程序数据类型。|
||[lastRefreshed](/javascript/api/excel/excel.linkeddatatype#lastrefreshed)|自上次刷新链接的工作簿时打开数据类型时区日期和时间。|
||[name](/javascript/api/excel/excel.linkeddatatype#name)|链接对象的名称数据类型。|
||[periodicRefreshInterval](/javascript/api/excel/excel.linkeddatatype#periodicrefreshinterval)|链接对象刷新的频率（以秒数据类型设置为 `refreshMode` "Periodic"时刷新。|
||[refreshMode](/javascript/api/excel/excel.linkeddatatype#refreshmode)|检索链接数据数据类型的机制。|
||[serviceId](/javascript/api/excel/excel.linkeddatatype#serviceid)|链接对象的唯一数据类型。|
||[supportedRefreshModes](/javascript/api/excel/excel.linkeddatatype#supportedrefreshmodes)|返回一个数组，该数组支持链接对象支持的所有数据类型。|
||[requestRefresh () ](/javascript/api/excel/excel.linkeddatatype#requestrefresh--)|请求刷新链接数据类型。|
||[requestSetRefreshMode (refreshMode： Excel.LinkedDataTypeRefreshMode) ](/javascript/api/excel/excel.linkeddatatype#requestsetrefreshmode-refreshmode-)|请求更改此链接的刷新数据类型。|
|[LinkedDataTypeAddedEventArgs](/javascript/api/excel/excel.linkeddatatypeaddedeventargs)|[serviceId](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#serviceid)|新链接对象的唯一数据类型。|
||[source](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#source)|获取事件源。|
||[type](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#type)|获取事件的类型。|
|[LinkedDataTypeCollection](/javascript/api/excel/excel.linkeddatatypecollection)|[getCount()](/javascript/api/excel/excel.linkeddatatypecollection#getcount--)|获取集合中链接的数据类型的数量。|
||[getItem (键：数字) ](/javascript/api/excel/excel.linkeddatatypecollection#getitem-key-)|按服务 ID 数据类型链接的标识符。|
||[getItemAt(index: number)](/javascript/api/excel/excel.linkeddatatypecollection#getitemat-index-)|按集合数据类型索引获取链接的索引。|
||[getItemOrNullObject (键：数字) ](/javascript/api/excel/excel.linkeddatatypecollection#getitemornullobject-key-)|按 ID 获取数据类型链接对象。|
||[items](/javascript/api/excel/excel.linkeddatatypecollection#items)|获取此集合中已加载的子项。|
||[requestRefreshAll () ](/javascript/api/excel/excel.linkeddatatypecollection#requestrefreshall--)|请求刷新集合中所有链接的数据类型。|
|[NamedSheetView](/javascript/api/excel/excel.namedsheetview)|[activate()](/javascript/api/excel/excel.namedsheetview#activate--)|激活此工作表视图。|
||[delete()](/javascript/api/excel/excel.namedsheetview#delete--)|从工作表中删除工作表视图。|
||[重复 (名称？：字符串) ](/javascript/api/excel/excel.namedsheetview#duplicate-name-)|创建此工作表视图的副本。|
||[name](/javascript/api/excel/excel.namedsheetview#name)|获取或设置工作表视图的名称。|
|[NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection)|[add(name: string)](/javascript/api/excel/excel.namedsheetviewcollection#add-name-)|创建具有给定名称的新工作表视图。|
||[enterTemporary () ](/javascript/api/excel/excel.namedsheetviewcollection#entertemporary--)|创建并激活新的临时工作表视图。|
||[exit () ](/javascript/api/excel/excel.namedsheetviewcollection#exit--)|退出当前活动的工作表视图。|
||[getActive () ](/javascript/api/excel/excel.namedsheetviewcollection#getactive--)|获取工作表当前的活动工作表视图。|
||[getCount()](/javascript/api/excel/excel.namedsheetviewcollection#getcount--)|获取此工作表中的工作表视图数。|
||[getItem(key: string)](/javascript/api/excel/excel.namedsheetviewcollection#getitem-key-)|使用工作表视图的名称获取工作表视图。|
||[getItemAt(index: number)](/javascript/api/excel/excel.namedsheetviewcollection#getitemat-index-)|按工作表视图在集合中的索引获取工作表视图。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.namedsheetviewcollection#getitemornullobject-key-)|使用工作表视图的名称获取工作表视图。|
||[items](/javascript/api/excel/excel.namedsheetviewcollection#items)|获取此集合中已加载的子项。|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[altTextDescription](/javascript/api/excel/excel.pivotlayout#alttextdescription)|数据透视表的替换文本说明。|
||[altTextTitle](/javascript/api/excel/excel.pivotlayout#alttexttitle)|数据透视表的替换文字标题。|
||[displayBlankLineAfterEachItem (显示： boolean) ](/javascript/api/excel/excel.pivotlayout#displayblanklineaftereachitem-display-)|设置是否在每个项后显示一个空行。|
||[emptyCellText](/javascript/api/excel/excel.pivotlayout#emptycelltext)|自动填充到数据透视表中任何空单元格的文本（如果 `fillEmptyCells == true` ）。|
||[fillEmptyCells](/javascript/api/excel/excel.pivotlayout#fillemptycells)|指定数据透视表中是否应当填充空单元格 `emptyCellText` 。|
||[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#getcell-datahierarchy--rowitems--columnitems-)|根据数据层次结构以及各自层次结构的行和列项，获取数据透视表中的唯一单元格。|
||[pivotStyle](/javascript/api/excel/excel.pivotlayout#pivotstyle)|应用于数据透视表的样式。|
||[repeatAllItemLabels (repeatLabels： boolean) ](/javascript/api/excel/excel.pivotlayout#repeatallitemlabels-repeatlabels-)|设置数据透视表中所有字段的"重复所有项目标签"设置。|
||[setStyle (：string \| PivotTableStyle \| BuiltInPivotTableStyle) ](/javascript/api/excel/excel.pivotlayout#setstyle-style-)|设置应用于数据透视表的样式。|
||[showFieldHeaders](/javascript/api/excel/excel.pivotlayout#showfieldheaders)|指定数据透视表是否显示字段标题 (字段标题和筛选下拉) 。|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[refreshOnOpen](/javascript/api/excel/excel.pivottable#refreshonopen)|指定工作簿打开时数据透视表是否刷新。|
|[PivotTableScopedCollection](/javascript/api/excel/excel.pivottablescopedcollection)|[getFirstOrNullObject () ](/javascript/api/excel/excel.pivottablescopedcollection#getfirstornullobject--)|获取集合中的第一个数据透视表。|
|[区域](/javascript/api/excel/excel.range)|[getDependents () ](/javascript/api/excel/excel.range#getdependents--)|返回一个对象，该对象代表包含同一工作表或多个工作表中单元格的所有 `WorkbookRangeAreas` 从属单元格的范围。|
||[getDirectDependents () ](/javascript/api/excel/excel.range#getdirectdependents--)|返回一个对象，该对象代表包含同一工作表或多个工作表中单元格的所有直接 `WorkbookRangeAreas` 从属单元格的范围。|
||[getExtendedRange (方向： Excel.KeyboardDirection， activeCell？： Range \| string) ](/javascript/api/excel/excel.range#getextendedrange-direction--activecell-)|返回一个 range 对象，该对象包含当前范围，并基于提供的方向返回区域边缘。|
||[getMergedAreasOrNullObject () ](/javascript/api/excel/excel.range#getmergedareasornullobject--)|返回一个 RangeAreas 对象，该对象代表此范围中的合并区域。|
||[getPrecedents () ](/javascript/api/excel/excel.range#getprecedents--)|返回一个对象，该对象代表包含同一工作表或多个工作表中单元格的所有引用 `WorkbookRangeAreas` 单元格的范围。|
||[getRangeEdge (方向：Excel.KeyboardDirection、activeCell？：range \| string) ](/javascript/api/excel/excel.range#getrangeedge-direction--activecell-)|返回一个 range 对象，该对象是数据区域的边缘单元格，对应于提供的方向。|
|[RefreshModeChangedEventArgs](/javascript/api/excel/excel.refreshmodechangedeventargs)|[refreshMode](/javascript/api/excel/excel.refreshmodechangedeventargs#refreshmode)|链接的数据类型刷新模式。|
||[serviceId](/javascript/api/excel/excel.refreshmodechangedeventargs#serviceid)|其刷新模式已更改的对象的唯一 ID。|
||[source](/javascript/api/excel/excel.refreshmodechangedeventargs#source)|获取事件源。|
||[type](/javascript/api/excel/excel.refreshmodechangedeventargs#type)|获取事件的类型。|
|[RefreshRequestCompletedEventArgs](/javascript/api/excel/excel.refreshrequestcompletedeventargs)|[已刷新](/javascript/api/excel/excel.refreshrequestcompletedeventargs#refreshed)|指示刷新请求是否成功。|
||[serviceId](/javascript/api/excel/excel.refreshrequestcompletedeventargs#serviceid)|已完成刷新请求的对象的唯一 ID。|
||[source](/javascript/api/excel/excel.refreshrequestcompletedeventargs#source)|获取事件源。|
||[type](/javascript/api/excel/excel.refreshrequestcompletedeventargs#type)|获取事件的类型。|
||[警告](/javascript/api/excel/excel.refreshrequestcompletedeventargs#warnings)|包含从刷新请求生成的任何警告的数组。|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addSvg(xml: string)](/javascript/api/excel/excel.shapecollection#addsvg-xml-)|从 XML 字符串创建可缩放的矢量图形 (SVG) 并将其添加到工作表。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.shapecollection#getitemornullobject-key-)|使用形状的名称或 ID 获取形状。|
|[Slicer](/javascript/api/excel/excel.slicer)|[nameInFormula](/javascript/api/excel/excel.slicer#nameinformula)|表示公式中使用切片器名称。|
||[slicerStyle](/javascript/api/excel/excel.slicer#slicerstyle)|应用于切片器的样式。|
||[setStyle (：string \| SlicerStyle \| BuiltInSlicerStyle) ](/javascript/api/excel/excel.slicer#setstyle-style-)|设置应用于切片器的样式。|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[getItemOrNullObject(name: string)](/javascript/api/excel/excel.stylecollection#getitemornullobject-name-)|按名称获取样式。|
|[Table](/javascript/api/excel/excel.table)|[clearStyle()](/javascript/api/excel/excel.table#clearstyle--)|将表格更改为使用默认表格样式。|
||[onFiltered](/javascript/api/excel/excel.table#onfiltered)|当筛选应用于特定表时发生。|
||[tableStyle](/javascript/api/excel/excel.table#tablestyle)|应用于表格的样式。|
||[resize (newRange： Range \| string) ](/javascript/api/excel/excel.table#resize-newrange-)|将表格调整为新区域。|
||[setStyle (样式：string \| TableStyle \| BuiltInTableStyle) ](/javascript/api/excel/excel.table#setstyle-style-)|设置应用于表格的样式。|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onFiltered](/javascript/api/excel/excel.tablecollection#onfiltered)|当对工作簿或工作表中的任意表应用筛选器时发生。|
|[TableFilteredEventArgs](/javascript/api/excel/excel.tablefilteredeventargs)|[tableId](/javascript/api/excel/excel.tablefilteredeventargs#tableid)|获取应用筛选器的表的 ID。|
||[type](/javascript/api/excel/excel.tablefilteredeventargs#type)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.tablefilteredeventargs#worksheetid)|获取包含表格的工作表的 ID。|
|[TableScopedCollection](/javascript/api/excel/excel.tablescopedcollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.tablescopedcollection#getitemornullobject-key-)|按名称或 ID 获取表。|
|[Workbook](/javascript/api/excel/excel.workbook)|[insertWorksheetsFromBase64 (base64File： string， options？： Excel.InsertWorksheetOptions) ](/javascript/api/excel/excel.workbook#insertworksheetsfrombase64-base64file--options-)|将源工作簿中的指定工作表插入到当前工作簿中。|
||[linkedDataTypes](/javascript/api/excel/excel.workbook#linkeddatatypes)|返回属于工作簿的链接数据类型的集合。|
||[onActivated](/javascript/api/excel/excel.workbook#onactivated)|在激活工作簿时发生。|
||[任务](/javascript/api/excel/excel.workbook#tasks)|返回工作簿中的任务的集合。|
||[showPivotFieldList](/javascript/api/excel/excel.workbook#showpivotfieldlist)|指定是否在工作簿级别显示数据透视表的字段列表窗格。|
||[use1904DateSystem](/javascript/api/excel/excel.workbook#use1904datesystem)|如果工作簿使用 1904 日期系统，则为 True。|
|[WorkbookActivatedEventArgs](/javascript/api/excel/excel.workbookactivatedeventargs)|[type](/javascript/api/excel/excel.workbookactivatedeventargs#type)|获取事件的类型。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[namedSheetViews](/javascript/api/excel/excel.worksheet#namedsheetviews)|返回工作表中呈现的工作表视图的集合。|
||[onFiltered](/javascript/api/excel/excel.worksheet#onfiltered)|当筛选应用于特定工作表时发生。|
||[onFormulaChanged](/javascript/api/excel/excel.worksheet#onformulachanged)|在此工作表中更改一个或多个公式时发生。|
||[任务](/javascript/api/excel/excel.worksheet#tasks)|返回工作表中的任务集合。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addfrombase64-base64file--sheetnamestoinsert--positiontype--relativeto-)|将工作簿的指定工作表插入当前工作簿。|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#onfiltered)|在工作簿中应用任何工作表的筛选器时发生。|
||[onFormulaChanged](/javascript/api/excel/excel.worksheetcollection#onformulachanged)|在此集合的任何工作表中更改一个或多个公式时发生。|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[type](/javascript/api/excel/excel.worksheetfilteredeventargs#type)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#worksheetid)|获取应用筛选器的工作表的 ID。|
|[WorksheetFormulaChangedEventArgs](/javascript/api/excel/excel.worksheetformulachangedeventargs)|[formulaDetails](/javascript/api/excel/excel.worksheetformulachangedeventargs#formuladetails)|获取对象 `FormulaChangedEventDetail` 数组，其中包含有关所有已更改公式的详细信息。|
||[source](/javascript/api/excel/excel.worksheetformulachangedeventargs#source)|事件的源。|
||[type](/javascript/api/excel/excel.worksheetformulachangedeventargs#type)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.worksheetformulachangedeventargs#worksheetid)|获取公式发生更改的工作表的 ID。|
