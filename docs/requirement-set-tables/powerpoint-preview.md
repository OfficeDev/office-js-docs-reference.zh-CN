| 类 | 域 | 说明 |
|:---|:---|:---|
|[AddSlideOptions](/javascript/api/powerpoint/powerpoint.addslideoptions)|[layoutId](/javascript/api/powerpoint/powerpoint.addslideoptions#layoutid)|指定要用于新幻灯片的幻灯片版式 ID。|
||[slideMasterId](/javascript/api/powerpoint/powerpoint.addslideoptions#slidemasterid)|指定要用于新幻灯片的幻灯片母版的 ID。|
|[Presentation](/javascript/api/powerpoint/powerpoint.presentation)|[slideMasters](/javascript/api/powerpoint/powerpoint.presentation#slidemasters)|返回演示文稿 `SlideMaster` 中的对象的集合。|
||[tags](/javascript/api/powerpoint/powerpoint.presentation#tags)|返回附加到演示文稿的标记的集合。|
|[Shape](/javascript/api/powerpoint/powerpoint.shape)|[id](/javascript/api/powerpoint/powerpoint.shape#id)|获取形状的唯一 ID。|
||[tags](/javascript/api/powerpoint/powerpoint.shape#tags)|返回形状中的标记集合。|
|[ShapeCollection](/javascript/api/powerpoint/powerpoint.shapecollection)|[getCount()](/javascript/api/powerpoint/powerpoint.shapecollection#getcount--)|获取集合中的形状数。|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.shapecollection#getitem-key-)|使用形状的唯一 ID 获取形状。|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.shapecollection#getitemat-index-)|使用形状在集合中从零开始编制的索引获取形状。|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.shapecollection#getitemornullobject-id-)|使用形状的唯一 ID 获取形状。|
||[items](/javascript/api/powerpoint/powerpoint.shapecollection#items)|获取此集合中已加载的子项。|
|[Slide](/javascript/api/powerpoint/powerpoint.slide)|[layout](/javascript/api/powerpoint/powerpoint.slide#layout)|获取幻灯片的版式。|
||[shapes](/javascript/api/powerpoint/powerpoint.slide#shapes)|返回幻灯片中形状的集合。|
||[slideMaster](/javascript/api/powerpoint/powerpoint.slide#slidemaster)|获取 `SlideMaster` 表示幻灯片的默认内容的对象。|
||[tags](/javascript/api/powerpoint/powerpoint.slide#tags)|返回幻灯片中的标记集合。|
|[SlideCollection](/javascript/api/powerpoint/powerpoint.slidecollection)|[添加 (选项？：PowerPoint.AddSlideOptions) ](/javascript/api/powerpoint/powerpoint.slidecollection#add-options-)|在集合的末尾添加新幻灯片。|
|[SlideLayout](/javascript/api/powerpoint/powerpoint.slidelayout)|[id](/javascript/api/powerpoint/powerpoint.slidelayout#id)|获取幻灯片版式的唯一 ID。|
||[名称](/javascript/api/powerpoint/powerpoint.slidelayout#name)|获取幻灯片版式的名称。|
|[SlideLayoutCollection](/javascript/api/powerpoint/powerpoint.slidelayoutcollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getcount--)|获取集合中的布局数。|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getitem-key-)|使用唯一 ID 获取布局。|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getitemat-index-)|获取一个布局，该布局使用集合中从零开始编制的索引。|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getitemornullobject-id-)|使用唯一 ID 获取布局。|
||[items](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#items)|获取此集合中已加载的子项。|
|[SlideMaster](/javascript/api/powerpoint/powerpoint.slidemaster)|[id](/javascript/api/powerpoint/powerpoint.slidemaster#id)|获取幻灯片母版的唯一 ID。|
||[布局](/javascript/api/powerpoint/powerpoint.slidemaster#layouts)|获取幻灯片母版提供的幻灯片版式的集合。|
||[名称](/javascript/api/powerpoint/powerpoint.slidemaster#name)|获取幻灯片母版的唯一名称。|
|[SlideMasterCollection](/javascript/api/powerpoint/powerpoint.slidemastercollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidemastercollection#getcount--)|获取集合中幻灯片母版的数量。|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidemastercollection#getitem-key-)|使用幻灯片母版的唯一 ID 获取幻灯片母版。|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidemastercollection#getitemat-index-)|使用集合中从零开始编制的索引获取幻灯片母版。|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.slidemastercollection#getitemornullobject-id-)|使用幻灯片母版的唯一 ID 获取幻灯片母版。|
||[items](/javascript/api/powerpoint/powerpoint.slidemastercollection#items)|获取此集合中已加载的子项。|
|[Tag](/javascript/api/powerpoint/powerpoint.tag)|[key](/javascript/api/powerpoint/powerpoint.tag#key)|获取标记的唯一 ID。|
||[value](/javascript/api/powerpoint/powerpoint.tag#value)|获取标记的值。|
|[TagCollection](/javascript/api/powerpoint/powerpoint.tagcollection)|[add (key： string， value： string) ](/javascript/api/powerpoint/powerpoint.tagcollection#add-key--value-)|在集合的末尾添加新标记。|
||[delete (key： string) ](/javascript/api/powerpoint/powerpoint.tagcollection#delete-key-)|删除此集合中给定 `key` 标记。|
||[getCount()](/javascript/api/powerpoint/powerpoint.tagcollection#getcount--)|获取集合中的标记数。|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.tagcollection#getitem-key-)|使用其唯一 ID 获取标记。|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.tagcollection#getitemat-index-)|获取一个标记，该标记使用集合中从零开始编制的索引。|
||[getItemOrNullObject(key: string)](/javascript/api/powerpoint/powerpoint.tagcollection#getitemornullobject-key-)|使用其唯一 ID 获取标记。|
||[items](/javascript/api/powerpoint/powerpoint.tagcollection#items)|获取此集合中已加载的子项。|
