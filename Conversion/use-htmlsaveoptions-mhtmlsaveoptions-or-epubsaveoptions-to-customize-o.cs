using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputPath = Path.Combine("Input", "Sample.docx");

        // -----------------------------------------------------------------
        // 1. Convert to HTML with custom image folder and export document properties.
        // -----------------------------------------------------------------
        Document htmlDoc = new Document(inputPath);

        // Create HtmlSaveOptions for HTML format.
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            // Export built‑in and custom document properties.
            ExportDocumentProperties = true,

            // Store linked images in a specific folder.
            ImagesFolder = Path.Combine("Output", "HtmlImages"),

            // Use UTF‑8 encoding without BOM.
            Encoding = new UTF8Encoding(false)
        };

        // Ensure the images folder exists.
        Directory.CreateDirectory(htmlOptions.ImagesFolder);

        // Save the document as HTML using the configured options.
        string htmlOutput = Path.Combine("Output", "Sample.html");
        htmlDoc.Save(htmlOutput, htmlOptions);

        // -----------------------------------------------------------------
        // 2. Convert to MHTML with CID URLs for resources and embed images as Base64.
        // -----------------------------------------------------------------
        Document mhtmlDoc = new Document(inputPath);

        // Create HtmlSaveOptions for MHTML format.
        HtmlSaveOptions mhtmlOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
        {
            // Use Content‑ID URLs for resources (images, fonts, CSS).
            ExportCidUrlsForMhtmlResources = true,

            // Embed images directly in the MHTML file using Base64.
            ExportImagesAsBase64 = true,

            // Export document properties if needed.
            ExportDocumentProperties = true
        };

        // Save the document as MHTML.
        string mhtmlOutput = Path.Combine("Output", "Sample.mhtml");
        mhtmlDoc.Save(mhtmlOutput, mhtmlOptions);

        // -----------------------------------------------------------------
        // 3. Convert to EPUB with split by heading paragraphs and UTF‑8 encoding.
        // -----------------------------------------------------------------
        Document epubDoc = new Document(inputPath);

        // Create HtmlSaveOptions for EPUB format.
        HtmlSaveOptions epubOptions = new HtmlSaveOptions(SaveFormat.Epub)
        {
            // Split the EPUB into separate HTML parts at each heading paragraph.
            DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph,

            // Export document properties into the EPUB package.
            ExportDocumentProperties = true,

            // Use UTF‑8 encoding.
            Encoding = Encoding.UTF8
        };

        // Save the document as EPUB.
        string epubOutput = Path.Combine("Output", "Sample.epub");
        epubDoc.Save(epubOutput, epubOptions);
    }
}
