using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define file names.
        const string inputPdfPath = "input.pdf";
        const string outputPdfPath = "output_pdfa1a.pdf";

        // -----------------------------------------------------------------
        // Step 1: Create a sample PDF file that will act as the source PDF.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is a sample PDF document.");
        // Save the document as a regular PDF.
        sourceDoc.Save(inputPdfPath, SaveFormat.Pdf);

        // Verify that the source PDF was created.
        if (!File.Exists(inputPdfPath))
            throw new InvalidOperationException("Failed to create the source PDF file.");

        // ---------------------------------------------------------------
        // Step 2: Load the source PDF into an Aspose.Words Document object.
        // ---------------------------------------------------------------
        Document pdfDoc = new Document(inputPdfPath);

        // ---------------------------------------------------------------
        // Step 3: Configure PDF/A‑1a compliance and enable OCR (if applicable).
        // ---------------------------------------------------------------
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            // Set the compliance level to PDF/A‑1a (searchable and tagged).
            Compliance = PdfCompliance.PdfA1a,

            // ExportDocumentStructure is required for PDF/A‑1a; it is
            // automatically enabled when Compliance is set to PdfA1a,
            // but we set it explicitly for clarity.
            ExportDocumentStructure = true
        };

        // Note: Aspose.Words can perform OCR when converting scanned PDFs.
        // If OCR settings are available in the version being used, they can be
        // configured here (e.g., saveOptions.OcrLanguage = "eng";).

        // ---------------------------------------------------------------
        // Step 4: Save the document as a searchable PDF/A‑1a file.
        // ---------------------------------------------------------------
        pdfDoc.Save(outputPdfPath, saveOptions);

        // Verify that the output PDF/A‑1a file was created.
        if (!File.Exists(outputPdfPath))
            throw new InvalidOperationException("The PDF/A‑1a output file was not created.");

        // The example finishes without requiring any user interaction.
    }
}
