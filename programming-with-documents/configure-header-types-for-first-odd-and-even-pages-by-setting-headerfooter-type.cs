using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a folder for output files.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Initialize a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Enable different headers/footers for the first page and for odd/even pages.
        builder.PageSetup.DifferentFirstPageHeaderFooter = true;
        builder.PageSetup.OddAndEvenPagesHeaderFooter = true;

        // ----- First page header -----
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderFirst);
        builder.Write("Header - First page");

        // ----- Even page header -----
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderEven);
        builder.Write("Header - Even pages");

        // ----- Primary (odd) page header -----
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Write("Header - Odd pages");

        // Return to the main document body.
        builder.MoveToSection(0);

        // Add three pages to demonstrate each header type.
        builder.Writeln("Content of page 1");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Content of page 2");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Content of page 3");

        // Save the resulting document.
        string outputPath = Path.Combine(artifactsDir, "HeadersAndFooters.docx");
        doc.Save(outputPath);
    }
}
