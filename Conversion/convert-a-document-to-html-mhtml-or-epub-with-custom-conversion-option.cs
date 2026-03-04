using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Saving;

class ConvertDocument
{
    static void Main()
    {
        // Load the source document from disk.
        Document doc = new Document("Input.docx");

        // -------------------- Convert to EPUB --------------------
        // Create HtmlSaveOptions for EPUB format and set custom options.
        HtmlSaveOptions epubOptions = new HtmlSaveOptions(SaveFormat.Epub)
        {
            // Use UTF-8 encoding without BOM.
            Encoding = new UTF8Encoding(false),

            // Export built‑in and custom document properties.
            ExportDocumentProperties = true,

            // Split the EPUB into separate HTML parts at heading paragraphs.
            DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph,

            // Store linked images in a separate folder (not embedded as Base64).
            ImagesFolder = "EpubImages",
            ExportImagesAsBase64 = false,
            ExportFontsAsBase64 = false
        };

        // Ensure the images folder exists.
        if (!Directory.Exists(epubOptions.ImagesFolder))
            Directory.CreateDirectory(epubOptions.ImagesFolder);

        // Save the document as EPUB using the configured options.
        doc.Save("Output.epub", epubOptions);

        // -------------------- Convert to HTML --------------------
        // Create HtmlSaveOptions for HTML format with different settings.
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            // Export form fields as plain text rather than HTML input elements.
            ExportTextInputFormFieldAsText = true,

            // Store images in a folder rather than embedding them.
            ImagesFolder = "HtmlImages",
            ExportImagesAsBase64 = false,

            // Produce nicely indented (pretty) HTML output.
            PrettyFormat = true
        };

        // Ensure the images folder exists.
        if (!Directory.Exists(htmlOptions.ImagesFolder))
            Directory.CreateDirectory(htmlOptions.ImagesFolder);

        // Save the document as HTML.
        doc.Save("Output.html", htmlOptions);

        // -------------------- Convert to MHTML --------------------
        // Create HtmlSaveOptions for MHTML format.
        HtmlSaveOptions mhtmlOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
        {
            // Use CID URLs for resources inside the MHTML package.
            ExportCidUrlsForMhtmlResources = true,

            // Embed images directly as Base64 within the MHTML.
            ExportImagesAsBase64 = true
        };

        // Save the document as MHTML.
        doc.Save("Output.mhtml", mhtmlOptions);
    }
}
