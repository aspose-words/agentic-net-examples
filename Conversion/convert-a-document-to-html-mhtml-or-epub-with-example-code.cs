using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsConversionExamples
{
    class Program
    {
        static void Main()
        {
            // Paths used in the examples – adjust as needed.
            string MyDir = @"C:\Input\";
            string ArtifactsDir = @"C:\Output\";

            // Ensure the output directory exists.
            if (!Directory.Exists(ArtifactsDir))
                Directory.CreateDirectory(ArtifactsDir);

            // Load a source document (DOCX, DOC, etc.).
            Document doc = new Document(Path.Combine(MyDir, "SampleDocument.docx"));

            // -------------------------------------------------
            // 1. Convert to HTML
            // -------------------------------------------------
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                // Export headers/footers per section (default) – can be changed to None if not needed.
                ExportHeadersFootersMode = ExportHeadersFootersMode.PerSection,

                // Store linked images in a separate folder.
                ImagesFolder = Path.Combine(ArtifactsDir, "HtmlImages"),

                // Optional: pretty‑print the HTML for readability.
                PrettyFormat = true
            };

            // Save the document as HTML.
            doc.Save(Path.Combine(ArtifactsDir, "SampleDocument.html"), htmlOptions);

            // -------------------------------------------------
            // 2. Convert to MHTML (Web archive)
            // -------------------------------------------------
            HtmlSaveOptions mhtmlOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
            {
                // Export document properties into the MHTML file.
                ExportDocumentProperties = true,

                // Use CID URLs for embedded resources (images, CSS, fonts).
                ExportCidUrlsForMhtmlResources = true
            };

            // Save the document as MHTML.
            doc.Save(Path.Combine(ArtifactsDir, "SampleDocument.mhtml"), mhtmlOptions);

            // -------------------------------------------------
            // 3. Convert to EPUB
            // -------------------------------------------------
            HtmlSaveOptions epubOptions = new HtmlSaveOptions(SaveFormat.Epub)
            {
                // Use UTF‑8 encoding without BOM.
                Encoding = new UTF8Encoding(false),

                // Split the EPUB into separate HTML parts at each heading paragraph.
                DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph,

                // Export built‑in and custom document properties.
                ExportDocumentProperties = true
            };

            // Save the document as EPUB.
            doc.Save(Path.Combine(ArtifactsDir, "SampleDocument.epub"), epubOptions);
        }
    }
}
