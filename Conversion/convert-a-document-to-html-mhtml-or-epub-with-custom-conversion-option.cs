using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsConversionDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source document.
            string sourcePath = @"C:\Docs\SourceDocument.docx";

            // Load the document using the built‑in constructor (rule: Document(string)).
            Document doc = new Document(sourcePath);

            // Convert to HTML with custom options.
            ConvertToHtml(doc, @"C:\Docs\Converted\Document.html");

            // Convert to MHTML with custom options.
            ConvertToMhtml(doc, @"C:\Docs\Converted\Document.mhtml");

            // Convert to EPUB with custom options.
            ConvertToEpub(doc, @"C:\Docs\Converted\Document.epub");
        }

        /// <summary>
        /// Saves the document as HTML using HtmlSaveOptions.
        /// </summary>
        static void ConvertToHtml(Document doc, string outputPath)
        {
            // Create HtmlSaveOptions for HTML format (rule: HtmlSaveOptions(SaveFormat)).
            HtmlSaveOptions options = new HtmlSaveOptions(SaveFormat.Html)
            {
                // Export built‑in and custom document properties.
                ExportDocumentProperties = true,

                // Use UTF‑8 encoding without BOM.
                Encoding = new UTF8Encoding(false),

                // Split the output into separate files at heading paragraphs.
                DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph,

                // Store linked images in a dedicated folder.
                ImagesFolder = Path.Combine(Path.GetDirectoryName(outputPath) ?? string.Empty, "Images")
            };

            // Ensure the images folder exists.
            if (!Directory.Exists(options.ImagesFolder))
                Directory.CreateDirectory(options.ImagesFolder);

            // Save the document using the overload that accepts SaveOptions (rule: Document.Save(string, SaveOptions)).
            doc.Save(outputPath, options);
        }

        /// <summary>
        /// Saves the document as MHTML using HtmlSaveOptions.
        /// </summary>
        static void ConvertToMhtml(Document doc, string outputPath)
        {
            // Create HtmlSaveOptions for MHTML format.
            HtmlSaveOptions options = new HtmlSaveOptions(SaveFormat.Mhtml)
            {
                ExportDocumentProperties = true,
                Encoding = new UTF8Encoding(false),

                // When saving MHTML, embed resources using CID URLs.
                ExportCidUrlsForMhtmlResources = true
            };

            doc.Save(outputPath, options);
        }

        /// <summary>
        /// Saves the document as EPUB using HtmlSaveOptions.
        /// </summary>
        static void ConvertToEpub(Document doc, string outputPath)
        {
            // Create HtmlSaveOptions for EPUB format.
            HtmlSaveOptions options = new HtmlSaveOptions(SaveFormat.Epub)
            {
                ExportDocumentProperties = true,
                Encoding = new UTF8Encoding(false),

                // Split the EPUB into separate HTML parts at heading paragraphs.
                DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph,

                // Export fonts as Base64 to ensure they are embedded.
                ExportFontsAsBase64 = true
            };

            doc.Save(outputPath, options);
        }
    }
}
