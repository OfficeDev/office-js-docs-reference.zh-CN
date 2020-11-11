| Class | 域 | 说明 |
|:---|:---|:---|
|[CustomXmlPart](/javascript/api/excel/excel.customxmlpart)|[delete()](/javascript/api/excel/excel.customxmlpart#delete--)|删除自定义 XML 部件。|
||[getXml ( # B1 ](/javascript/api/excel/excel.customxmlpart#getxml--)|获取自定义 XML 部件的完整 XML 内容。|
||[id](/javascript/api/excel/excel.customxmlpart#id)|自定义 XML 部件的 ID。|
||[namespaceUri](/javascript/api/excel/excel.customxmlpart#namespaceuri)|自定义 XML 部件的命名空间 URI。|
||[setXml (xml： string) ](/javascript/api/excel/excel.customxmlpart#setxml-xml-)|设置自定义 XML 部件的完整 XML 内容。|
|[CustomXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|[add (xml： string) ](/javascript/api/excel/excel.customxmlpartcollection#add-xml-)|向工作簿添加新的自定义 XML 部件。|
||[getByNamespace (namespaceUri： string) ](/javascript/api/excel/excel.customxmlpartcollection#getbynamespace-namespaceuri-)|获取其命名空间匹配给定命名空间的自定义 XML 部件的新作用域内集合。|
||[getCount()](/javascript/api/excel/excel.customxmlpartcollection#getcount--)|获取此集合中 CustomXml 部件的数量。|
||[getItem(id: string)](/javascript/api/excel/excel.customxmlpartcollection#getitem-id-)|获取基于其 ID 的自定义 XML 部件。|
||[getItemOrNullObject(id: string)](/javascript/api/excel/excel.customxmlpartcollection#getitemornullobject-id-)|获取基于其 ID 的自定义 XML 部件。|
||[items](/javascript/api/excel/excel.customxmlpartcollection#items)|获取此集合中已加载的子项。|
|[CustomXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|[getCount()](/javascript/api/excel/excel.customxmlpartscopedcollection#getcount--)|获取此集合中 CustomXML 部件的数量。|
||[getItem(id: string)](/javascript/api/excel/excel.customxmlpartscopedcollection#getitem-id-)|获取基于其 ID 的自定义 XML 部件。|
||[getItemOrNullObject(id: string)](/javascript/api/excel/excel.customxmlpartscopedcollection#getitemornullobject-id-)|获取基于其 ID 的自定义 XML 部件。|
||[getOnlyItem ( # B1 ](/javascript/api/excel/excel.customxmlpartscopedcollection#getonlyitem--)|如果集合仅包含一个项，则此方法返回该项。|
||[getOnlyItemOrNullObject ( # B1 ](/javascript/api/excel/excel.customxmlpartscopedcollection#getonlyitemornullobject--)|如果集合仅包含一个项，则此方法返回该项。|
||[items](/javascript/api/excel/excel.customxmlpartscopedcollection#items)|获取此集合中已加载的子项。|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[id](/javascript/api/excel/excel.pivottable#id)|数据透视表的 ID。|
|[RequestContext](/javascript/api/excel/excel.requestcontext)|[runtime](/javascript/api/excel/excel.requestcontext#runtime)|[Api set： ExcelApi 1.5]|
|[运行时](/javascript/api/excel/excel.runtime)||[Workbook](/javascript/api/excel/excel.workbook)|[customXmlParts](/javascript/api/excel/excel.workbook#customxmlparts)|表示此工作簿包含的自定义 XML 部件的集合。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[getNext (visibleOnly？： boolean) ](/javascript/api/excel/excel.worksheet#getnext-visibleonly-)|获取此工作表的后面的工作表。|
||[getNextOrNullObject (visibleOnly？： boolean) ](/javascript/api/excel/excel.worksheet#getnextornullobject-visibleonly-)|获取此工作表的后面的工作表。|
||[getPrevious (visibleOnly？： boolean) ](/javascript/api/excel/excel.worksheet#getprevious-visibleonly-)|获取此项之前的工作表。|
||[getPreviousOrNullObject (visibleOnly？： boolean) ](/javascript/api/excel/excel.worksheet#getpreviousornullobject-visibleonly-)|获取此项之前的工作表。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[getFirst (visibleOnly？： boolean) ](/javascript/api/excel/excel.worksheetcollection#getfirst-visibleonly-)|获取集合中的第一个工作表。|
||[getLast (visibleOnly？： boolean) ](/javascript/api/excel/excel.worksheetcollection#getlast-visibleonly-)|获取集合中的最后一个工作表。|
