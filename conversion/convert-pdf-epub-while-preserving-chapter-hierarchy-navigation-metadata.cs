using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Saving;

namespace PdfToEpubConversion
{
    class Program
    {
        static void Main()
        {
            // Path to the source PDF file.
            const string pdfPath = @"C:\Input\SourceDocument.pdf";

            // Path where the resulting EPUB file will be saved.
            const string epubPath = @"C:\Output\ResultDocument.epub";

            // Load the PDF document into an Aspose.Words Document object.
            // The constructor automatically detects the format from the file extension.
            Document pdfDocument = new Document(pdfPath);

            // Configure save options for EPUB output.
            HtmlSaveOptions epubSaveOptions = new HtmlSaveOptions
            {
                // Specify that the target format is EPUB.
                SaveFormat = SaveFormat.Epub,

                // Use UTF‑8 encoding without BOM.
                Encoding = new UTF8Encoding(false),

                // Split the document at heading paragraphs so that each chapter becomes a separate HTML part
                // inside the EPUB package. This preserves the chapter hierarchy.
                DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph,

                // Export built‑in and custom document properties (e.g., title, author) to the EPUB.
                ExportDocumentProperties = true,

                // Limit the navigation map (table of contents) to the desired heading levels.
                // Adjust this value according to the depth of your chapter hierarchy (1‑9).
                NavigationMapLevel = 3
            };

            // Save the loaded PDF as an EPUB using the configured options.
            pdfDocument.Save(epubPath, epubSaveOptions);
        }
    }
}
