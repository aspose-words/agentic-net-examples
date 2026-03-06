using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsSaveOptionsDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source DOCX file.
            string sourceDocPath = Path.Combine("MyDir", "SampleDocument.docx");

            // Load the DOCX document.
            Document doc = new Document(sourceDocPath);

            // -----------------------------------------------------------------
            // 1. Save as HTML with custom options.
            // -----------------------------------------------------------------
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                // Export built‑in and custom document properties.
                ExportDocumentProperties = true,

                // Split the output into separate HTML files at each heading paragraph.
                DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph,

                // Use UTF‑8 encoding without BOM.
                Encoding = new UTF8Encoding(false),

                // Store linked images in a dedicated folder.
                ImagesFolder = Path.Combine("ArtifactsDir", "HtmlImages"),

                // Export images as Base64 strings embedded in the HTML.
                ExportImagesAsBase64 = true,

                // Pretty‑format the generated HTML for readability.
                PrettyFormat = true
            };

            // Ensure the images folder exists.
            Directory.CreateDirectory(htmlOptions.ImagesFolder);

            // Save the document as HTML using the configured options.
            doc.Save(Path.Combine("ArtifactsDir", "SampleDocument.html"), htmlOptions);

            // -----------------------------------------------------------------
            // 2. Save as MHTML with custom options.
            // -----------------------------------------------------------------
            HtmlSaveOptions mhtmlOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
            {
                // Export document properties.
                ExportDocumentProperties = true,

                // Use CID URLs for embedded resources (images, fonts, CSS).
                ExportCidUrlsForMhtmlResources = true,

                // Set the encoding to UTF‑8.
                Encoding = Encoding.UTF8,

                // Do not split the document (MHTML does not support splitting).
                DocumentSplitCriteria = DocumentSplitCriteria.None,

                // Pretty‑format the output.
                PrettyFormat = true
            };

            // Save the document as MHTML.
            doc.Save(Path.Combine("ArtifactsDir", "SampleDocument.mhtml"), mhtmlOptions);

            // -----------------------------------------------------------------
            // 3. Save as EPUB with custom options.
            // -----------------------------------------------------------------
            HtmlSaveOptions epubOptions = new HtmlSaveOptions(SaveFormat.Epub)
            {
                // Export document properties into the EPUB package.
                ExportDocumentProperties = true,

                // Split the EPUB into multiple HTML parts at each heading paragraph.
                DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph,

                // Use UTF‑8 encoding.
                Encoding = Encoding.UTF8,

                // Export fonts as Base64 to embed them in the EPUB.
                ExportFontsAsBase64 = true,

                // Export images as Base64 to embed them in the EPUB.
                ExportImagesAsBase64 = true,

                // Pretty‑format the generated HTML parts.
                PrettyFormat = true
            };

            // Save the document as EPUB.
            doc.Save(Path.Combine("ArtifactsDir", "SampleDocument.epub"), epubOptions);
        }
    }
}
