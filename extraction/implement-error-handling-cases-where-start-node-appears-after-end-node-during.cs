using System;
using Aspose.Words;
using Aspose.Words.Layout;

namespace AsposeWordsExtraction
{
    class Program
    {
        static void Main()
        {
            // -----------------------------------------------------------------
            // Create a sample document with two bookmarks: "Start" and "End".
            // This makes the example self‑contained and avoids missing file errors.
            // -----------------------------------------------------------------
            Document srcDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(srcDoc);

            // Page 1 – contains the start bookmark.
            builder.Writeln("This is page 1.");
            builder.StartBookmark("Start");
            builder.Writeln("Start of the range.");
            builder.EndBookmark("Start");

            // Insert a page break to create a second page.
            builder.InsertBreak(BreakType.PageBreak);

            // Page 2 – contains the end bookmark.
            builder.Writeln("This is page 2.");
            builder.StartBookmark("End");
            builder.Writeln("End of the range.");
            builder.EndBookmark("End");

            // Insert another page break and some extra content.
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("This is page 3 – outside the range.");

            // Ensure the layout information is up‑to‑date.
            srcDoc.UpdatePageLayout();

            // Retrieve the start and end nodes using the bookmarks.
            Node startNode = srcDoc.Range.Bookmarks["Start"]?.BookmarkStart;
            Node endNode = srcDoc.Range.Bookmarks["End"]?.BookmarkStart;

            if (startNode == null || endNode == null)
            {
                Console.WriteLine("Start or end bookmark not found.");
                return;
            }

            // Map nodes to page numbers.
            LayoutCollector layout = new LayoutCollector(srcDoc);
            int startPage = layout.GetStartPageIndex(startNode);
            int endPage = layout.GetEndPageIndex(endNode);

            // Validate that both nodes could be mapped to pages.
            if (startPage == 0 || endPage == 0)
            {
                Console.WriteLine("One of the nodes cannot be mapped to a page.");
                return;
            }

            // Validate logical order of the nodes.
            if (startPage > endPage)
            {
                Console.WriteLine("The start node appears after the end node. Extraction aborted.");
                return;
            }

            // Convert to zero‑based index required by ExtractPages.
            int zeroBasedStartIndex = startPage - 1;
            int pageCount = endPage - startPage + 1;

            // Extract the required page range.
            Document extractedDoc = srcDoc.ExtractPages(zeroBasedStartIndex, pageCount);

            // Save the extracted document.
            extractedDoc.Save("Extracted.docx");
            Console.WriteLine($"Extracted pages {startPage}-{endPage} to 'Extracted.docx'.");
        }
    }
}
