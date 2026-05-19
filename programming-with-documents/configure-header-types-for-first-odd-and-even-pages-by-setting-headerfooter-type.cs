using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Enable different headers/footers for the first page and for odd/even pages.
        builder.PageSetup.DifferentFirstPageHeaderFooter = true;
        builder.PageSetup.OddAndEvenPagesHeaderFooter = true;

        // First page header.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderFirst);
        builder.Write("Header for the first page");

        // Even page header.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderEven);
        builder.Write("Header for even pages");

        // Primary (odd) page header.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Write("Header for odd pages");

        // Add three pages to demonstrate the different headers.
        builder.MoveToSection(0);
        builder.Writeln("Page 1");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 2");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 3");

        // Optional: output the header types to the console.
        foreach (HeaderFooter hf in doc.FirstSection.HeadersFooters)
        {
            Console.WriteLine($"Header/Footer type: {hf.HeaderFooterType}");
        }

        // Save the document to an output folder.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "HeadersAndFooters.docx");
        doc.Save(outputPath);
    }
}
