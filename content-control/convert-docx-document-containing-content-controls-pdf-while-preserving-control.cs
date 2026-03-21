using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Markup;

namespace ContentControlToPdf
{
    class Program
    {
        static void Main()
        {
            // Define paths relative to the current working directory.
            string sourceDocxPath = Path.Combine(Directory.GetCurrentDirectory(), "DocumentWithContentControls.docx");
            string outputPdfPath = Path.Combine(Directory.GetCurrentDirectory(), "DocumentWithContentControls.pdf");

            // Create a new DOCX document with a content control (StructuredDocumentTag).
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a plain‑text content control.
            StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Block);
            sdt.Title = "SampleControl";
            sdt.Tag = "SampleTag";

            // Append the content control directly to the document body.
            doc.FirstSection.Body.AppendChild(sdt);

            // Move the builder into the content control and write text.
            builder.MoveTo(sdt);
            builder.Write("This is sample text inside a content control.");

            // Save the DOCX so it can be re‑loaded if needed.
            doc.Save(sourceDocxPath);

            // Load the DOCX document (demonstrating the typical workflow).
            Document loadedDoc = new Document(sourceDocxPath);

            // Configure PDF save options to preserve form fields.
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                PreserveFormFields = true,
                UseSdtTagAsFormFieldName = true
            };

            // Save the document as PDF using the configured options.
            loadedDoc.Save(outputPdfPath, pdfOptions);

            Console.WriteLine($"DOCX saved to: {sourceDocxPath}");
            Console.WriteLine($"PDF saved to: {outputPdfPath}");
        }
    }
}
