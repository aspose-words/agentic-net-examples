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
        const string outputPdfPath = "output_searchable_pdfa1a.pdf";

        // -----------------------------------------------------------------
        // Step 1: Create a sample PDF file.
        // In a real scenario the PDF would already exist; here we generate one.
        // -----------------------------------------------------------------
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("This is a sample PDF created by Aspose.Words.");
        // Save as PDF.
        sampleDoc.Save(inputPdfPath, SaveFormat.Pdf);

        // -----------------------------------------------------------------
        // Step 2: Load the existing PDF.
        // -----------------------------------------------------------------
        Document pdfDocument = new Document(inputPdfPath);

        // -----------------------------------------------------------------
        // Step 3: Configure PDF save options.
        // Set compliance to PDF/A‑1a which requires a searchable document.
        // Aspose.Words will embed the document structure; OCR is applied
        // automatically when converting to PDF/A‑1a if the source contains
        // raster text.
        // -----------------------------------------------------------------
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA1a
        };

        // -----------------------------------------------------------------
        // Step 4: Save the PDF as a searchable PDF/A‑1a document.
        // -----------------------------------------------------------------
        pdfDocument.Save(outputPdfPath, saveOptions);

        // -----------------------------------------------------------------
        // Step 5: Verify that the output file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(outputPdfPath))
            throw new InvalidOperationException($"The file '{outputPdfPath}' was not created.");

        // Optional: clean up the temporary input file.
        if (File.Exists(inputPdfPath))
            File.Delete(inputPdfPath);
    }
}
