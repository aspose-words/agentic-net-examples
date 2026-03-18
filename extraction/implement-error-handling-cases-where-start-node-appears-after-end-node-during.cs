using System;
using Aspose.Words;
using Aspose.Words.Layout;

namespace AsposeWordsExtraction
{
    class Program
    {
        static void Main()
        {
            // Load the source document.
            Document srcDoc = new Document("Input.docx");

            // Ensure the layout information is up‑to‑date.
            srcDoc.UpdatePageLayout();

            // Retrieve the start and end nodes. In this example we use bookmarks named "Start" and "End".
            // Adjust the retrieval logic as needed for your scenario.
            Node startNode = srcDoc.Range.Bookmarks["Start"]?.BookmarkStart;
            Node endNode = srcDoc.Range.Bookmarks["End"]?.BookmarkStart;

            if (startNode == null || endNode == null)
                throw new InvalidOperationException("Start or end bookmark not found.");

            // Create a LayoutCollector to map nodes to page numbers.
            LayoutCollector layout = new LayoutCollector(srcDoc);

            // Get the 1‑based page numbers where the nodes begin and end.
            int startPage = layout.GetStartPageIndex(startNode);
            int endPage = layout.GetEndPageIndex(endNode);

            // Validate that both nodes could be mapped to pages.
            if (startPage == 0 || endPage == 0)
                throw new InvalidOperationException("One of the nodes cannot be mapped to a page.");

            // Validate the logical order of the nodes.
            if (startPage > endPage)
                throw new ArgumentException("The start node appears after the end node. Extraction aborted.");

            // Convert to zero‑based index required by ExtractPages.
            int zeroBasedStartIndex = startPage - 1;
            int pageCount = endPage - startPage + 1;

            // Extract the required page range.
            Document extractedDoc = srcDoc.ExtractPages(zeroBasedStartIndex, pageCount);

            // Save the extracted document.
            extractedDoc.Save("Extracted.docx");
        }
    }
}
