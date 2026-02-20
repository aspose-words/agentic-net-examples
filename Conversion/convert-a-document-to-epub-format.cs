using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source document (any supported format).
        Document doc = new Document("input.docx");

        // Create save options for EPUB using the HtmlSaveOptions constructor that accepts a SaveFormat.
        HtmlSaveOptions epubOptions = new HtmlSaveOptions(SaveFormat.Epub);

        // Optional: customize EPUB output.
        epubOptions.ExportDocumentProperties = true;          // Include built‑in and custom properties.
        epubOptions.ExportImagesAsBase64 = true;             // Embed images directly in the EPUB.
        epubOptions.ExportFontsAsBase64 = true;              // Embed fonts directly in the EPUB.
        epubOptions.ExportGeneratorName = true;              // Add Aspose.Words generator info.
        epubOptions.ExportHeadersFootersMode = ExportHeadersFootersMode.None; // Omit headers/footers.

        // Save the document as an EPUB file using the configured options.
        doc.Save("output.epub", epubOptions);
    }
}
