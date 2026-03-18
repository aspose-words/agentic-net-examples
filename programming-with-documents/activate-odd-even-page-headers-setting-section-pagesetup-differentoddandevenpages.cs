using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Activate different headers/footers for odd and even pages.
        builder.PageSetup.OddAndEvenPagesHeaderFooter = true;

        // Define the header that will appear on odd-numbered pages (primary header).
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Write("Header for odd pages");

        // Define the header that will appear on even-numbered pages.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderEven);
        builder.Write("Header for even pages");

        // Add some content and page breaks to create multiple pages.
        builder.MoveToSection(0);
        builder.Writeln("Content of page 1");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Content of page 2");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Content of page 3");

        // Save the document to disk.
        doc.Save("OddEvenHeaders.docx");
    }
}
