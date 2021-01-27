| 类 | 域 | 说明 |
|:---|:---|:---|
|[InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions)|[格式设置](/javascript/api/powerpoint/powerpoint.insertslideoptions#formatting)|指定在幻灯片插入过程中使用的格式。|
||[sourceSlideIds](/javascript/api/powerpoint/powerpoint.insertslideoptions#sourceslideids)|指定源演示文稿中将插入到当前演示文稿中的幻灯片。|
||[targetSlideId](/javascript/api/powerpoint/powerpoint.insertslideoptions#targetslideid)|指定演示文稿中新幻灯片的插入位置。|
|[Presentation](/javascript/api/powerpoint/powerpoint.presentation)|[insertSlidesFromBase64 (base64File： string， options？： PowerPoint.InsertSlideOptions) ](/javascript/api/powerpoint/powerpoint.presentation#insertslidesfrombase64-base64file--options-)|将演示文稿中的指定幻灯片插入到当前演示文稿中。|
||[幻灯片](/javascript/api/powerpoint/powerpoint.presentation#slides)|返回演示文稿中幻灯片的有序集合。|
|[Slide](/javascript/api/powerpoint/powerpoint.slide)|[delete()](/javascript/api/powerpoint/powerpoint.slide#delete--)|从演示文稿中删除幻灯片。|
||[id](/javascript/api/powerpoint/powerpoint.slide#id)|获取幻灯片的唯一 ID。|
|[SlideCollection](/javascript/api/powerpoint/powerpoint.slidecollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidecollection#getcount--)|获取集合中的幻灯片数。|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidecollection#getitem-key-)|使用唯一 ID 获取幻灯片。|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidecollection#getitemat-index-)|使用集合中从零开始索引获取幻灯片。|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.slidecollection#getitemornullobject-id-)|使用唯一 ID 获取幻灯片。|
||[items](/javascript/api/powerpoint/powerpoint.slidecollection#items)|获取此集合中已加载的子项。|
