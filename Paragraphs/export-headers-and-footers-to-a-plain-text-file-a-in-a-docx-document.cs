using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ExportHeadersFooters
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add a primary header with some text.
        HeaderFooter header = new HeaderFooter(doc, HeaderFooterType.HeaderPrimary);
        doc.FirstSection.HeadersFooters.Add(header);
        header.AppendParagraph("Primary header");

        // Add a primary footer with some text.
        HeaderFooter footer = new HeaderFooter(doc, HeaderFooterType.FooterPrimary);
        doc.FirstSection.HeadersFooters.Add(footer);
        footer.AppendParagraph("Primary footer");

        // Insert body content.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Body line 1");
        builder.Writeln("Body line 2");

        // Configure save options to include headers and footers in the plain‑text output.
        TxtSaveOptions saveOptions = new TxtSaveOptions
        {
            // Options: PrimaryOnly, AllAtEnd, or None.
            ExportHeadersFootersMode = TxtExportHeadersFootersMode.PrimaryOnly
        };

        // Save the document as a .txt file using the specified options.
        doc.Save("Exported.txt", saveOptions);
    }
}
