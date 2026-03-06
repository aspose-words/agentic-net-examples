using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ExportHeadersFootersToTxt
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        doc.EnsureMinimum(); // Guarantees at least one section, body, and paragraph.

        // ----- Add a primary header -----
        HeaderFooter primaryHeader = new HeaderFooter(doc, HeaderFooterType.HeaderPrimary);
        primaryHeader.AppendParagraph("Primary header");
        doc.FirstSection.HeadersFooters.Add(primaryHeader);

        // ----- Add a primary footer -----
        HeaderFooter primaryFooter = new HeaderFooter(doc, HeaderFooterType.FooterPrimary);
        primaryFooter.AppendParagraph("Primary footer");
        doc.FirstSection.HeadersFooters.Add(primaryFooter);

        // ----- Insert body content with page breaks -----
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Page 1");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 2");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Write("Page 3");

        // ----- Configure TXT save options to include headers/footers -----
        TxtSaveOptions txtOptions = new TxtSaveOptions();
        // Options: None, PrimaryOnly, AllAtEnd
        txtOptions.ExportHeadersFootersMode = TxtExportHeadersFootersMode.PrimaryOnly;

        // ----- Save the document as plain text -----
        doc.Save("HeadersFooters.txt", txtOptions);
    }
}
