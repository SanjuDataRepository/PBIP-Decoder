# Code-Store
Overview
This project is a utility designed to analyze Power BI projects (specifically .pbip format). It extracts and maps metadata across Bookmarks, Pages, and Visuals using a combination of Python and Java. The final output is an Excel file containing consolidated tabs for detailed report analysis.

**Prerequisites**

File Format: The Power BI file must be saved as a .pbip (Power BI Project) format.

Note: Do not use .pbix or .pbir files directly.

**Code Structure**
The project consists of the following modules/readers:

- Bookmark Reader
- Page Reader
- Bookmark Page Reader
- Log File Reader
- Pbip Log Reader (Primary Module)

**Workflow & Algorithm**
The process involves a hybrid approach using Python for file parsing and Java/Browser Console for visual extraction.

1. Python reads the Bookmarks.
2. Python reads the Pages.
3. Java extracts Page IDs from the playground environment.
4. Java searches for the extracted Page IDs and identifies the visuals within those pages from the playground.
5. Open the browser console and drop down the logs generated in Step 4 regarding the visuals.
6. Save these browser logs as a .log file.
7. Python reads the saved visual .log file.
8. Add Page Name into the Bookmarks tab (sourced from Visuals).
9. Add Visual Title into both the Bookmarks and Page tabs (sourced from Visuals).
10. Return a final Excel file containing three tabs: Bookmarks, Pages, and Visuals.

```text
+-----------------------+          +-----------------------+          +-----------------------+
|       BOOKMARKS       |          |         PAGE          |          |        VISUALS        |
|     (PBIP Folders)    |          |     (PBIP Folders)    |          |    (PBI Playground)   |
+-----------------------+          +-----------------------+          +-----------------------+
| Bookmark ID           |<-------->| Page ID               |          | Page ID               |
| Bookmark Name         | ID Match | Page Name             |          | Page Name (Visual)    |
| Visual ID             |          | Visual ID             |          | Type (Visual Type)    |
| Visual Type           |          | Visual Type           |          | Title (Visual Title)  |
| Mode                  |          | Action Type           |          +-----------------------+
| Selected Visual       |          |                       |                      |
| Applied Filters       |          |                       |                      |
| Slicer Selections     |          |                       |                      |
|                       |          |                       |                      |
| Visual Title <--------|----------|--- Visual Title <-----|----------------------+
| Page Name <-----------|----------|-----------------------+
+-----------------------+          +-----------------------+
```

 Mappings:
 - ID Match: Links Bookmark ID to Page ID.
 - Visual Title: Extracted from Visuals populated in Page and Bookmarks.
 - Page Name: Extracted from Visuals populated in Bookmarks.
