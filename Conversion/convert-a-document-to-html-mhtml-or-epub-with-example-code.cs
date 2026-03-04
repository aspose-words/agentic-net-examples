using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsConversionExamples
{
    class Program
    {
        // Adjust these paths to point to your input and output folders.
        private const string MyDir = @"C:\Input\";
        private const string ArtifactsDir = @"C:\Output\";

        static void Main()
        {
            // Load a source document (DOCX in this example).
            Document doc = new Document(Path.Combine(MyDir, "SampleDocument.docx"));

            // -----------------------------------------------------------------
            // 1. Convert to HTML
            // -----------------------------------------------------------------
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                // Export headers/footers per section (default) – can be changed if needed.
                ExportHeadersFootersMode = ExportHeadersFootersMode.PerSection,

                // Save images to a dedicated folder.
                ImagesFolder = Path.Combine(ArtifactsDir, "HtmlImages"),

                // Optional: pretty‑print the HTML for readability.
                PrettyFormat = true
            };

            // Ensure the images folder exists.
            if (Directory.Exists(htmlOptions.ImagesFolder) == false)
                Directory.CreateDirectory(htmlOptions.ImagesFolder);

            // Save the document as HTML.
            doc.Save(Path.Combine(ArtifactsDir, "SampleDocument.html"), htmlOptions);

            // -----------------------------------------------------------------
            // 2. Convert to MHTML (Web archive)
            // -----------------------------------------------------------------
            HtmlSaveOptions mhtmlOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
            {
                // Export CID URLs for embedded resources – required for MHTML.
                ExportCidUrlsForMhtmlResources = true,

                // Optional: embed images as Base64 to keep everything inside the MHTML file.
                ExportImagesAsBase64 = true
            };

            // Save the document as MHTML.
            doc.Save(Path.Combine(ArtifactsDir, "SampleDocument.mhtml"), mhtmlOptions);

            // -----------------------------------------------------------------
            // 3. Convert to EPUB
            // -----------------------------------------------------------------
            HtmlSaveOptions epubOptions = new HtmlSaveOptions(SaveFormat.Epub)
            {
                // Use UTF‑8 encoding (recommended for EPUB).
                Encoding = Encoding.UTF8,

                // Split the EPUB into separate HTML parts at each heading paragraph.
                DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph,

                // Export document properties (author, title, etc.) into the EPUB metadata.
                ExportDocumentProperties = true
            };

            // Save the document as EPUB.
            doc.Save(Path.Combine(ArtifactsDir, "SampleDocument.epub"), epubOptions);
        }
    }
}
