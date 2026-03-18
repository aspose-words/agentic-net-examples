using System;
using Aspose.Words;
using Aspose.Words.Saving;

class HeaderFooterIndexer
{
    static void Main()
    {
        // Load an existing Word document (load rule)
        Document doc = new Document("Input.docx");

        // Get the primary header of the first section and read its plain text via the Range object
        HeaderFooter header = doc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary];
        string headerText = header?.Range?.Text ?? string.Empty;

        // Get the primary footer of the first section and read its plain text via the Range object
        HeaderFooter footer = doc.FirstSection.HeadersFooters[HeaderFooterType.FooterPrimary];
        string footerText = footer?.Range?.Text ?? string.Empty;

        // Combine header and footer texts for indexing (example)
        string indexContent = $"Header: {headerText.Trim()}\nFooter: {footerText.Trim()}";

        // Output the combined text (could be fed to an indexing engine)
        Console.WriteLine(indexContent);

        // Save a plain‑text version of the document that includes the primary header/footer
        TxtSaveOptions txtOptions = new TxtSaveOptions
        {
            ExportHeadersFootersMode = TxtExportHeadersFootersMode.PrimaryOnly
        };
        doc.Save("Output.txt", txtOptions);
    }
}
