using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source document (replace with your actual file path).
        Document doc = new Document("Input.docx");

        // Create a SaveOptions object suitable for PDF format.
        // The factory method returns a SaveOptions instance; cast it to PdfSaveOptions.
        SaveOptions saveOptions = SaveOptions.CreateSaveOptions(SaveFormat.Pdf);
        PdfSaveOptions pdfOptions = (PdfSaveOptions)saveOptions;

        // Set the PDF/A compliance level (e.g., PDF/A-1b) to ensure archival quality.
        pdfOptions.Compliance = PdfCompliance.PdfA1b;

        // Export the document structure, which is required for PDF/A compliance.
        pdfOptions.ExportDocumentStructure = true;

        // Save the document as a PDF file using the configured options.
        doc.Save("Output.pdf", pdfOptions);
    }
}
