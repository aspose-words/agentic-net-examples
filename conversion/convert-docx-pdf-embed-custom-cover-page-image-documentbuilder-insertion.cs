using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsExample
{
    class Program
    {
        static void Main()
        {
            // Path to the source DOCX file.
            string sourceDocx = @"C:\Docs\SourceDocument.docx";

            // Path to the custom cover page image.
            string coverImage = @"C:\Images\CoverPage.jpg";

            // Path where the resulting PDF will be saved.
            string outputPdf = @"C:\Docs\ResultDocument.pdf";

            // Load the existing DOCX document.
            Document doc = new Document(sourceDocx);

            // Create a DocumentBuilder attached to the loaded document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Move the cursor to the very beginning of the document.
            builder.MoveToDocumentStart();

            // Insert the cover page image. The image is inserted inline at 100% scale.
            builder.InsertImage(coverImage);

            // Save the modified document as PDF.
            doc.Save(outputPdf, SaveFormat.Pdf);
        }
    }
}
