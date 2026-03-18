using Aspose.Words;
using Aspose.Words.Layout;
using Aspose.Words.Settings;

class HyphenationPaginationDemo
{
    static void Main()
    {
        // Load the multi‑section report document.
        Document doc = new Document("Input.docx");

        // Build the initial layout and obtain the page count.
        doc.UpdatePageLayout();
        LayoutCollector layout = new LayoutCollector(doc);
        int pagesBefore = layout.GetNumPagesSpanned(doc);

        // Enable automatic hyphenation and configure optional parameters.
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.HyphenationOptions.HyphenationZone = 720;          // 0.5 inch from the right margin.
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;    // Limit consecutive hyphenated lines.

        // Re‑build the layout after changing hyphenation settings.
        doc.UpdatePageLayout();

        // Clear previous layout data and obtain the new page count.
        layout.Clear();
        int pagesAfter = layout.GetNumPagesSpanned(doc);

        // Save the hyphenated document for verification.
        doc.Save("Output_Hyphenated.docx");

        // Calculate and display the pagination difference.
        int pageDifference = pagesAfter - pagesBefore;
        System.Console.WriteLine($"Pages before hyphenation: {pagesBefore}");
        System.Console.WriteLine($"Pages after hyphenation: {pagesAfter}");
        System.Console.WriteLine($"Difference: {pageDifference}");
    }
}
