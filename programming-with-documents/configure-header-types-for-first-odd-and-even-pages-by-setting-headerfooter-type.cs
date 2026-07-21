using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder for convenience.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Enable different headers for the first page and for odd/even pages.
        builder.PageSetup.DifferentFirstPageHeaderFooter = true;
        builder.PageSetup.OddAndEvenPagesHeaderFooter = true;

        // ----- First page header -----
        HeaderFooter firstHeader = new HeaderFooter(doc, HeaderFooterType.HeaderFirst);
        firstHeader.AppendParagraph("Header for the first page");
        doc.FirstSection.HeadersFooters.Add(firstHeader);

        // ----- Even page header -----
        HeaderFooter evenHeader = new HeaderFooter(doc, HeaderFooterType.HeaderEven);
        evenHeader.AppendParagraph("Header for even pages");
        doc.FirstSection.HeadersFooters.Add(evenHeader);

        // ----- Primary (odd) page header -----
        HeaderFooter primaryHeader = new HeaderFooter(doc, HeaderFooterType.HeaderPrimary);
        primaryHeader.AppendParagraph("Header for odd pages");
        doc.FirstSection.HeadersFooters.Add(primaryHeader);

        // Add some content to generate multiple pages.
        for (int i = 1; i <= 3; i++)
        {
            builder.Writeln($"Page {i}");
            if (i < 3)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(outputDir);

        // Save the document.
        string outputPath = Path.Combine(outputDir, "HeadersExample.docx");
        doc.Save(outputPath);
    }
}
