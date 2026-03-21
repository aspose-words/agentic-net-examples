using Aspose.Words;
using Aspose.Words.Layout;

class HyphenationPaginationDemo
{
    static void Main()
    {
        // Create a new document with a long paragraph to generate multiple pages.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Size = 12;
        string longText = string.Concat(
            "Lorem ipsum dolor sit amet, consectetur adipiscing elit. ",
            "Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. ",
            "Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. ",
            "Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. ",
            "Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum. "
        );
        // Repeat the text to ensure the document spans several pages.
        for (int i = 0; i < 50; i++)
        {
            builder.Writeln(longText);
        }

        // Build the initial layout and obtain the page count.
        doc.UpdatePageLayout();
        LayoutCollector layout = new LayoutCollector(doc);
        int pagesBefore = layout.GetNumPagesSpanned(doc.FirstSection.Body);

        // Enable automatic hyphenation and configure optional parameters.
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.HyphenationOptions.HyphenationZone = 720;          // 0.5 inch from the right margin.
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;    // Limit consecutive hyphenated lines.

        // Re‑build the layout after changing hyphenation settings.
        doc.UpdatePageLayout();

        // Clear previous layout data and obtain the new page count.
        layout.Clear();
        int pagesAfter = layout.GetNumPagesSpanned(doc.FirstSection.Body);

        // Save the hyphenated document for verification.
        doc.Save("Output_Hyphenated.docx");

        // Calculate and display the pagination difference.
        int pageDifference = pagesAfter - pagesBefore;
        System.Console.WriteLine($"Pages before hyphenation: {pagesBefore}");
        System.Console.WriteLine($"Pages after hyphenation: {pagesAfter}");
        System.Console.WriteLine($"Difference: {pageDifference}");
    }
}
