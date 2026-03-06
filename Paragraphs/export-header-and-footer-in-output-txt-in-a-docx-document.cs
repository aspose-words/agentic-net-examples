using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class ExportHeadersFootersToTxt
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
        builder.Writeln("Page 1");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 2");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Write("Page 3");

        // Configure save options to export headers and footers.
        TxtSaveOptions saveOptions = new TxtSaveOptions
        {
            // Export only primary headers/footers at the start/end of each section.
            ExportHeadersFootersMode = TxtExportHeadersFootersMode.PrimaryOnly
        };

        // Define the output TXT file path.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "ExportHeadersFooters.txt");

        // Save the document as plain text using the configured options.
        doc.Save(outputPath, saveOptions);
    }
}
