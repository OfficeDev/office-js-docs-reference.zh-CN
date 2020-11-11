| Class | 域 | 说明 |
|:---|:---|:---|
|[InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions)|[编排](/javascript/api/powerpoint/powerpoint.insertslideoptions#formatting)|指定在幻灯片插入过程中要使用的格式。|
||[sourceSlideIds](/javascript/api/powerpoint/powerpoint.insertslideoptions#sourceslideids)|指定将插入到当前演示文稿中的源演示文稿中的幻灯片。|
||[targetSlideId](/javascript/api/powerpoint/powerpoint.insertslideoptions#targetslideid)|指定将在演示文稿中插入新幻灯片的位置。|
|[Presentation](/javascript/api/powerpoint/powerpoint.presentation)|[insertSlidesFromBase64 (base64File： string，options？： InsertSlideOptions) ](/javascript/api/powerpoint/powerpoint.presentation#insertslidesfrombase64-base64file--options-)|将演示文稿中指定的幻灯片插入到当前演示文稿中。|
||[页面](/javascript/api/powerpoint/powerpoint.presentation#slides)|返回演示文稿中的幻灯片的已排序集合。|
|[Slide](/javascript/api/powerpoint/powerpoint.slide)|[delete()](/javascript/api/powerpoint/powerpoint.slide#delete--)|删除演示文稿中的幻灯片。|
||[id](/javascript/api/powerpoint/powerpoint.slide#id)|获取幻灯片的唯一 ID。|
|[SlideCollection](/javascript/api/powerpoint/powerpoint.slidecollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidecollection#getcount--)|获取集合中的幻灯片数。|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidecollection#getitem-key-)|使用其唯一 ID 获取幻灯片。|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidecollection#getitemat-index-)|使用集合中的从零开始的索引获取幻灯片。|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.slidecollection#getitemornullobject-id-)|使用其唯一 ID 获取幻灯片。|
||[items](/javascript/api/powerpoint/powerpoint.slidecollection#items)|获取此集合中已加载的子项。|
