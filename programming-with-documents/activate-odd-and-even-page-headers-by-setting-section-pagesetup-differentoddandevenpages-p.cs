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

        // Create an even‑page header.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderEven);
        builder.Write("Even Page Header");

        // Create an odd‑page (primary) header.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Write("Odd Page Header");

        // Return to the main body of the first section.
        builder.MoveToSection(0);
        builder.Writeln("Page 1");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 2");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 3");

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Save the document.
        string outputPath = Path.Combine(outputDir, "OddEvenHeaders.docx");
        doc.Save(outputPath);
    }
}
