using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source document (replace with your actual file path).
        Document doc = new Document("Input.docx");

        // Create a PdfSaveOptions instance via the factory method (lifecycle rule).
        PdfSaveOptions pdfOptions = SaveOptions.CreateSaveOptions(SaveFormat.Pdf) as PdfSaveOptions;

        // Specify the required PDF compliance level.
        // Change the enum value to the compliance you need (e.g., PdfA1b, PdfUa1, Pdf20, etc.).
        pdfOptions.Compliance = PdfCompliance.PdfA1b;

        // Save the document as PDF using the configured options.
        doc.Save("Output.pdf", pdfOptions);
    }
}
