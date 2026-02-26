using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Use DocumentBuilder to add content with proper heading styles.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add a heading (will be exported as a tag in the PDF structure).
            builder.ParagraphFormat.Style = doc.Styles["Heading 1"];
            builder.Writeln("Document Title");

            // Add a normal paragraph.
            builder.ParagraphFormat.Style = doc.Styles["Normal"];
            builder.Writeln("This is a sample paragraph that will appear in the PDF document.");

            // Prepare PDF save options with document structure export enabled.
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                ExportDocumentStructure = true // Preserve tags for headings, paragraphs, etc.
            };

            // Save the document as PDF with the specified options.
            doc.Save("OutputWithStructure.pdf", pdfOptions);
        }
    }
}
