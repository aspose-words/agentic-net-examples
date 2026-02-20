using System;
using Aspose.Words;
using Aspose.Words.Layout;

class Program
{
    static void Main()
    {
        // Load the DOCX document.
        Document doc = new Document("Input.docx");

        // Force layout to be calculated so that page numbers are available.
        // This is required before we can query page information.
        doc.UpdatePageLayout();

        // Attach a LayoutCollector to the document.
        // The collector maps document nodes to their page numbers.
        LayoutCollector collector = new LayoutCollector(doc);

        // Iterate through each page in the document.
        // The Document class does not expose a direct PageCollection,
        // but we can use the PageCount property to enumerate pages.
        for (int pageIndex = 1; pageIndex <= doc.PageCount; pageIndex++)
        {
            // Example: retrieve the first node that starts on the current page.
            // We walk through the document nodes and find one whose start page matches the loop index.
            Node nodeOnPage = FindFirstNodeOnPage(doc, collector, pageIndex);

            // Output basic information about the page.
            Console.WriteLine($"--- Page {pageIndex} ---");
            if (nodeOnPage != null)
            {
                Console.WriteLine($"First node on page: {nodeOnPage.GetType().Name}");
                Console.WriteLine($"Text snippet: {nodeOnPage.GetText().Trim().Substring(0, Math.Min(30, nodeOnPage.GetText().Trim().Length))}");
            }
            else
            {
                Console.WriteLine("No node found on this page.");
            }
        }
    }

    // Helper method that scans the document's child nodes and returns the first node whose
    // start page index matches the supplied page number.
    private static Node FindFirstNodeOnPage(Document doc, LayoutCollector collector, int pageNumber)
    {
        foreach (Node node in doc.GetChildNodes(NodeType.Any, true))
        {
            int startPage = collector.GetStartPageIndex(node);
            if (startPage == pageNumber)
                return node;
        }
        return null;
    }
}
