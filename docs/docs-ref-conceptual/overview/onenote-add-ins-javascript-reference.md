# <a name="onenote-javascript-api-overview"></a><span data-ttu-id="b9841-101">OneNote JavaScript API 概述</span><span class="sxs-lookup"><span data-stu-id="b9841-101">OneNote JavaScript API overview</span></span>

<span data-ttu-id="b9841-102">适用于：OneNote Online</span><span class="sxs-lookup"><span data-stu-id="b9841-102">Applies to: OneNote Online</span></span>

<span data-ttu-id="b9841-103">下面的链接展示了 API 中的高级 OneNote 对象。</span><span class="sxs-lookup"><span data-stu-id="b9841-103">The following links show the high level OneNote objects available in the API.</span></span> <span data-ttu-id="b9841-104">每个对象页上链接包含的属性、 事件和方法的对象上的说明。</span><span class="sxs-lookup"><span data-stu-id="b9841-104">Each object page link contains a description of the properties, events, and methods available on the object.</span></span> <span data-ttu-id="b9841-105">如需了解详细信息，请浏览相应链接。</span><span class="sxs-lookup"><span data-stu-id="b9841-105">Explore these links to learn more.</span></span> 
    
- <span data-ttu-id="b9841-106">[Application](/javascript/api/onenote/onenote.application)：用于访问所有全局可寻址的 OneNote 对象（如活动笔记本和活动分区）的顶级对象。</span><span class="sxs-lookup"><span data-stu-id="b9841-106">[Application](/javascript/api/onenote/onenote.application): The top-level object used to access all globally addressable OneNote objects, such as the active notebook and the active section.</span></span>

- <span data-ttu-id="b9841-p102">[笔记本](/javascript/api/onenote/onenote.notebook)：一个笔记本。笔记本包含分区组合和分区。</span><span class="sxs-lookup"><span data-stu-id="b9841-p102">[Notebook](/javascript/api/onenote/onenote.notebook): A notebook. Notebooks contain section groups and sections.</span></span>
    - <span data-ttu-id="b9841-109">[NotebookCollection](/javascript/api/onenote/onenote.notebookcollection)：笔记本的集合。</span><span class="sxs-lookup"><span data-stu-id="b9841-109">[NotebookCollection](/javascript/api/onenote/onenote.notebookcollection): A collection of notebooks.</span></span>

- <span data-ttu-id="b9841-p103">[SectionGroup](/javascript/api/onenote/onenote.sectiongroup)：一个分区组。分区组包含分区组和分区。</span><span class="sxs-lookup"><span data-stu-id="b9841-p103">[SectionGroup](/javascript/api/onenote/onenote.sectiongroup): A section group. Section groups contain section groups and sections.</span></span>
    - <span data-ttu-id="b9841-112">[SectionGroupCollection](/javascript/api/onenote/onenote.sectiongroupcollection)：分区组的集合。</span><span class="sxs-lookup"><span data-stu-id="b9841-112">[SectionGroupCollection](/javascript/api/onenote/onenote.sectiongroupcollection): A collection of section groups.</span></span>

- <span data-ttu-id="b9841-p104">[Section](/javascript/api/onenote/onenote.section)：一个分区。分区包含页面。</span><span class="sxs-lookup"><span data-stu-id="b9841-p104">[Section](/javascript/api/onenote/onenote.section): A section. Sections contain pages.</span></span>
    - <span data-ttu-id="b9841-115">[SectionCollection](/javascript/api/onenote/onenote.sectioncollection)：分区的集合。</span><span class="sxs-lookup"><span data-stu-id="b9841-115">[SectionCollection](/javascript/api/onenote/onenote.sectioncollection): A collection of sections.</span></span>

- <span data-ttu-id="b9841-p105">[Page](/javascript/api/onenote/onenote.page)：一个页面。页面包含 PageContent 对象。</span><span class="sxs-lookup"><span data-stu-id="b9841-p105">[Page](/javascript/api/onenote/onenote.page): A page. Pages contain PageContent objects.</span></span>
    - <span data-ttu-id="b9841-118">[PageCollection](/javascript/api/onenote/onenote.pagecollection)：页面的集合。</span><span class="sxs-lookup"><span data-stu-id="b9841-118">[PageCollection](/javascript/api/onenote/onenote.pagecollection): A collection of pages.</span></span>

- <span data-ttu-id="b9841-p106">[PageContent](/javascript/api/onenote/onenote.pagecontent)：页面上包含内容类型的顶级地区，例如 Outline 或 Image。可在页面上为 PageContent 对象分配一个位置。</span><span class="sxs-lookup"><span data-stu-id="b9841-p106">[PageContent](/javascript/api/onenote/onenote.pagecontent): A top-level region on a page that contains content types such as Outline or Image. A PageContent object can be assigned a position on the page.</span></span>
    - <span data-ttu-id="b9841-121">[PageContentCollection](/javascript/api/onenote/onenote.pagecontentcollection)：PageContent 对象的集合，表示页面的内容。</span><span class="sxs-lookup"><span data-stu-id="b9841-121">[PageContentCollection](/javascript/api/onenote/onenote.pagecontentcollection): A collection of PageContent objects, which represents the contents of a page.</span></span>

- <span data-ttu-id="b9841-p107">[Outline](/javascript/api/onenote/onenote.outline)：Paragraph 对象的容器。Outline 是 PageContent 对象的直接子级。</span><span class="sxs-lookup"><span data-stu-id="b9841-p107">[Outline](/javascript/api/onenote/onenote.outline): A container for Paragraph objects. An Outline is a direct child of a PageContent object.</span></span>

- <span data-ttu-id="b9841-p108">[Image](/javascript/api/onenote/onenote.image)：Image 对象。Image 可以是 PageContent 对象或 Paragraph 的直接子级。</span><span class="sxs-lookup"><span data-stu-id="b9841-p108">[Image](/javascript/api/onenote/onenote.image): An Image object. An Image can be a direct child of a PageContent object or a Paragraph.</span></span>

- <span data-ttu-id="b9841-p109">[Paragraph](/javascript/api/onenote/onenote.paragraph)：页面上可见内容的容器。Paragraph 是 Outline 的直接子级。</span><span class="sxs-lookup"><span data-stu-id="b9841-p109">[Paragraph](/javascript/api/onenote/onenote.paragraph): A container for the visible content on a page. A Paragraph is a direct child of an Outline.</span></span>
    - <span data-ttu-id="b9841-128">[ParagraphCollection](/javascript/api/onenote/onenote.paragraphcollection)：Outline 中 Paragraph 对象的集合。</span><span class="sxs-lookup"><span data-stu-id="b9841-128">[ParagraphCollection](/javascript/api/onenote/onenote.paragraphcollection): A collection of Paragraph objects in an Outline.</span></span>

- <span data-ttu-id="b9841-129">[RichText](/javascript/api/onenote/onenote.richtext)：RichText 对象。</span><span class="sxs-lookup"><span data-stu-id="b9841-129">[RichText](/javascript/api/onenote/onenote.richtext): A RichText object.</span></span>

- <span data-ttu-id="b9841-130">[表格](/javascript/api/onenote/onenote.table)：TableRow 对象的容器。</span><span class="sxs-lookup"><span data-stu-id="b9841-130">[Table](/javascript/api/onenote/onenote.table): A container for TableRow objects.</span></span>

- <span data-ttu-id="b9841-131">[TableRow](/javascript/api/onenote/onenote.tablerow)：TableCell 对象的容器。</span><span class="sxs-lookup"><span data-stu-id="b9841-131">[TableRow](/javascript/api/onenote/onenote.tablerow): A container for TableCell objects.</span></span>
    - <span data-ttu-id="b9841-132">[TableRowCollection](/javascript/api/onenote/onenote.tablerowcollection)：表中 TableRow 对象的集合。</span><span class="sxs-lookup"><span data-stu-id="b9841-132">[TableRowCollection](/javascript/api/onenote/onenote.tablerowcollection): A collection of TableRow objects in a Table.</span></span>
 
- <span data-ttu-id="b9841-133">[TableCell](/javascript/api/onenote/onenote.tablecell)：段落对象的容器。</span><span class="sxs-lookup"><span data-stu-id="b9841-133">[TableCell](/javascript/api/onenote/onenote.tablecell): A container for Paragraph objects.</span></span>
    - <span data-ttu-id="b9841-134">[TableCellCollection](/javascript/api/onenote/onenote.tablecellcollection)：TableRow 中 TableCell 对象的集合。</span><span class="sxs-lookup"><span data-stu-id="b9841-134">[TableCellCollection](/javascript/api/onenote/onenote.tablecellcollection): A collection of TableCell objects in a TableRow.</span></span>

## <a name="onenote-javascript-api-reference"></a><span data-ttu-id="b9841-135">OneNote JavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="b9841-135">OneNote JavaScript API reference</span></span>

<span data-ttu-id="b9841-136">有关 OneNote JavaScript API 的详细信息，请参阅[OneNote JavaScript API 参考文档](/javascript/api/onenote)。</span><span class="sxs-lookup"><span data-stu-id="b9841-136">For detailed information about OneNote JavaScript API, see the [OneNote JavaScript API reference documentation](/javascript/api/onenote).</span></span>

## <a name="see-also"></a><span data-ttu-id="b9841-137">另请参阅</span><span class="sxs-lookup"><span data-stu-id="b9841-137">See also</span></span>

- [<span data-ttu-id="b9841-138">OneNote JavaScript API 编程概述</span><span class="sxs-lookup"><span data-stu-id="b9841-138">OneNote JavaScript API programming overview</span></span>](https://docs.microsoft.com/office/dev/add-ins/onenote/onenote-add-ins-programming-overview)
- [<span data-ttu-id="b9841-139">生成第一个 OneNote 外接程序</span><span class="sxs-lookup"><span data-stu-id="b9841-139">Build your first OneNote add-in</span></span>](https://docs.microsoft.com/office/dev/add-ins/onenote/onenote-add-ins-getting-started)
- [<span data-ttu-id="b9841-140">Rubric Grader 示例</span><span class="sxs-lookup"><span data-stu-id="b9841-140">Rubric Grader sample</span></span>](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [<span data-ttu-id="b9841-141">Office 加载项平台概述</span><span class="sxs-lookup"><span data-stu-id="b9841-141">Office Add-ins platform overview</span></span>](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-ins)
