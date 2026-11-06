using System;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        Section section = doc.FirstSection;

        // Enable different headers for first page and for odd/even pages.
        section.PageSetup.DifferentFirstPageHeaderFooter = true;
        section.PageSetup.OddAndEvenPagesHeaderFooter = true;

        // Create and add a header for the first page.
        HeaderFooter firstPageHeader = new HeaderFooter(doc, HeaderFooterType.HeaderFirst);
        section.HeadersFooters.Add(firstPageHeader);
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderFirst);
        builder.Writeln("First Page Header");

        // Create and add a header for odd pages (primary header).
        HeaderFooter oddHeader = new HeaderFooter(doc, HeaderFooterType.HeaderPrimary);
        section.HeadersFooters.Add(oddHeader);
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Writeln("Odd Page Header");

        // Create and add a header for even pages.
        HeaderFooter evenHeader = new HeaderFooter(doc, HeaderFooterType.HeaderEven);
        section.HeadersFooters.Add(evenHeader);
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderEven);
        builder.Writeln("Even Page Header");

        // Add body content to generate multiple pages.
        builder.MoveToDocumentEnd();
        for (int i = 1; i <= 5; i++)
        {
            builder.Writeln($"Paragraph {i}");
            builder.InsertBreak(BreakType.PageBreak);
        }

        // Save the document.
        string outputPath = "HeaderTypes.docx";
        doc.Save(outputPath);
    }
}
