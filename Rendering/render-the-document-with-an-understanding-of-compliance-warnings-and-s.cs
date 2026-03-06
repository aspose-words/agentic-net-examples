using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("Input.docx");

        // Configure PDF save options to use a specific compliance level.
        // Here we choose PDF/A-1b which preserves visual appearance.
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA1b
        };

        // Save the document as a PDF with the defined compliance settings.
        doc.Save("Output.pdf", saveOptions);
    }
}
