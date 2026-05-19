using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Prepare output folder and file name.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "OddEvenHeaders.docx");

        // Create a new blank document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Enable different headers/footers for odd and even pages.
        builder.PageSetup.OddAndEvenPagesHeaderFooter = true;

        // Header for odd-numbered pages (primary header).
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Write("Header for odd pages");

        // Header for even-numbered pages.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderEven);
        builder.Write("Header for even pages");

        // Return to the main document body and add some pages.
        builder.MoveToSection(0);
        builder.Writeln("Page 1 (odd)");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 2 (even)");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 3 (odd)");

        // Save the document.
        doc.Save(outputPath);
    }
}
