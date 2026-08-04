using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Enable different headers/footers for odd and even pages.
        // This setting affects all sections in the document.
        builder.PageSetup.OddAndEvenPagesHeaderFooter = true;

        // Create an odd (primary) header.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Write("Header for odd pages");

        // Create an even header.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderEven);
        builder.Write("Header for even pages");

        // Return to the main body of the first section.
        builder.MoveToSection(0);

        // Add three pages to demonstrate odd/even headers.
        builder.Writeln("Page 1 (odd)");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 2 (even)");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 3 (odd)");

        // Ensure the output directory exists.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Save the document.
        string outputPath = Path.Combine(artifactsDir, "OddEvenHeaders.docx");
        doc.Save(outputPath);
    }
}
