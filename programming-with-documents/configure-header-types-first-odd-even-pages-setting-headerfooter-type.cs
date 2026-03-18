using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Enable distinct headers for the first page and for odd/even pages.
        // These settings affect the whole section (the first and only section at this point).
        doc.FirstSection.PageSetup.DifferentFirstPageHeaderFooter = true;
        doc.FirstSection.PageSetup.OddAndEvenPagesHeaderFooter = true;

        // ----- First page header -----
        // Create a HeaderFooter of type HeaderFirst and add a paragraph to it.
        HeaderFooter headerFirst = new HeaderFooter(doc, HeaderFooterType.HeaderFirst);
        headerFirst.AppendParagraph("Header – First page");
        // Add the header to the section's HeadersFooters collection.
        doc.FirstSection.HeadersFooters.Add(headerFirst);

        // ----- Odd (primary) page header -----
        // HeaderPrimary is used for odd-numbered pages when OddAndEvenPagesHeaderFooter is true.
        HeaderFooter headerOdd = new HeaderFooter(doc, HeaderFooterType.HeaderPrimary);
        headerOdd.AppendParagraph("Header – Odd pages");
        doc.FirstSection.HeadersFooters.Add(headerOdd);

        // ----- Even page header -----
        // HeaderEven is used for even-numbered pages.
        HeaderFooter headerEven = new HeaderFooter(doc, HeaderFooterType.HeaderEven);
        headerEven.AppendParagraph("Header – Even pages");
        doc.FirstSection.HeadersFooters.Add(headerEven);

        // Add body content that spans multiple pages to demonstrate the headers.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Page 1");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 2");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 3");

        // Save the document to disk.
        doc.Save("HeadersConfigured.docx");
    }
}
