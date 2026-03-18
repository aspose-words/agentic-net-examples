using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace ContentControlToPdf
{
    class Program
    {
        static void Main()
        {
            // Path to the source DOCX that contains content controls (StructuredDocumentTag).
            string sourceDocxPath = @"C:\Input\DocumentWithContentControls.docx";

            // Path where the resulting PDF will be saved.
            string outputPdfPath = @"C:\Output\DocumentWithContentControls.pdf";

            // Load the DOCX document.
            Document doc = new Document(sourceDocxPath);

            // Configure PDF save options.
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Preserve Word content controls as interactive PDF form fields.
                PreserveFormFields = true,

                // Use the SDT tag (or Id) as the name of the PDF form field.
                // This helps to keep a clear mapping between the original control and the PDF field.
                UseSdtTagAsFormFieldName = true
            };

            // Save the document as PDF using the configured options.
            doc.Save(outputPdfPath, pdfOptions);
        }
    }
}
