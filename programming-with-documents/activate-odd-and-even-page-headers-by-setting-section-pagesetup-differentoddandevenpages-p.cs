using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Enable different headers for odd and even pages.
        builder.PageSetup.OddAndEvenPagesHeaderFooter = true;

        // Create a header that will appear on odd-numbered pages.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Writeln("Header for odd pages");

        // Create a header that will appear on even-numbered pages.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderEven);
        builder.Writeln("Header for even pages");

        // Return to the main document body.
        builder.MoveToSection(0);

        // Add a few pages to demonstrate the odd/even headers.
        builder.Writeln("Page 1");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 2");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 3");

        // Determine an output path in the current working directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "OddEvenHeaders.docx");
        doc.Save(outputPath);
    }
}
