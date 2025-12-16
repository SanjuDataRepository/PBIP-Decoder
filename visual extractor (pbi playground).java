// Retrieve the page collection and get the visuals for multiple pages.
try {
    const pages = await report.getPages();
    // Supply multiple internal page IDs (these correspond to page.name)
    const pageIds = [
'Insert page id',
'Insert page id',
];
    // Index pages by id for quick lookup
    const pagesById = new Map(pages.map(p => [p.name, p]));
    for (const pageId of pageIds) {
        // Find the page by its internal name (ID)
        const page = pagesById.get(pageId);
        if (!page) {
            console.warn(`Page with ID "${pageId}" not found.`);
            continue;
        }
        const visuals = await page.getVisuals();
        console.log({
            pageId,
            visuals: visuals.map(function (visual) {
                return {
                    pageId: pageId,
                    name: visual.name,
                    type: visual.type,
                    title: visual.title,
                    layout: visual.layout
                };
            })
        });
    }
}
catch (errors) {
    console.log(errors);
}
