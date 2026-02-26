using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source document (replace with your actual file path).
        Document doc = new Document("Input.docx");

        // Create a save options object suitable for PDF using the provided factory method.
        SaveOptions saveOptions = SaveOptions.CreateSaveOptions(SaveFormat.Pdf);

        // Cast to PdfSaveOptions to configure PDF‑specific settings.
        PdfSaveOptions pdfOptions = (PdfSaveOptions)saveOptions;

        // Choose a PDF/A compliance level (e.g., PDF/A‑2u).
        pdfOptions.Compliance = PdfCompliance.PdfA2u;

        // Save the document as a PDF that conforms to the selected PDF/A standard.
        doc.Save("Output.pdf", pdfOptions);
    }
}
