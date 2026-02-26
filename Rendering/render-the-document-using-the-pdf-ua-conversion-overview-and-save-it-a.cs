using System;
using Aspose.Words;
using Aspose.Words.Saving;

class PdfUaConversion
{
    static void Main()
    {
        // Load the source Word document.
        // Replace "Input.docx" with the actual path to your document.
        Document doc = new Document("Input.docx");

        // Create a PdfSaveOptions object suitable for PDF output.
        // The CreateSaveOptions method ensures we use the correct lifecycle rule.
        SaveOptions saveOptions = SaveOptions.CreateSaveOptions(SaveFormat.Pdf);

        // Cast to PdfSaveOptions to access PDF‑specific properties.
        PdfSaveOptions pdfOptions = (PdfSaveOptions)saveOptions;

        // Set the compliance level to PDF/UA‑2 (ISO 14289‑2:2024).
        pdfOptions.Compliance = PdfCompliance.PdfUa2;

        // PDF/UA requires the document title to be shown in the viewer's title bar.
        pdfOptions.DisplayDocTitle = true;

        // Optional: export the document structure (tags) – required for PDF/UA compliance.
        pdfOptions.ExportDocumentStructure = true;

        // Save the document as a PDF with the specified options.
        // Replace "Output.pdf" with the desired output path.
        doc.Save("Output.pdf", pdfOptions);
    }
}
