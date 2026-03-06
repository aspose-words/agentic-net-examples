using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source Word document.
        Document doc = new Document("Input.docx");

        // Create a PdfSaveOptions instance to customize PDF output.
        PdfSaveOptions saveOptions = new PdfSaveOptions();

        // Set the PDF compliance level.
        // Options include PdfCompliance.PdfA1a, PdfCompliance.PdfA1b, PdfCompliance.PdfA2u, etc.
        // Here we use PDF/A-1b as an example.
        saveOptions.Compliance = PdfCompliance.PdfA1b;

        // Save the document as a PDF file using the specified options.
        doc.Save("Output.pdf", saveOptions);
    }
}
