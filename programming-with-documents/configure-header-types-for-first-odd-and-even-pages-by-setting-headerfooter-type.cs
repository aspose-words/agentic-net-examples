using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Define output folder and file.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(artifactsDir);
        string outputPath = Path.Combine(artifactsDir, "HeadersAndFooters.docx");

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Enable different headers for first page and for odd/even pages.
        builder.PageSetup.DifferentFirstPageHeaderFooter = true;
        builder.PageSetup.OddAndEvenPagesHeaderFooter = true;

        // Create the first‑page header.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderFirst);
        builder.Write("Header for the first page");

        // Create the even‑page header.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderEven);
        builder.Write("Header for even pages");

        // Create the primary (odd‑page) header.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Write("Header for odd pages");

        // Return to the main document body and add three pages to demonstrate the headers.
        builder.MoveToSection(0);
        builder.Writeln("Page 1 (first page)");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 2 (even page)");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 3 (odd page)");

        // Save the document.
        doc.Save(outputPath);
    }
}
