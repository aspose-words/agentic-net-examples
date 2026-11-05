using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Drawing;

public class PdfAConversionExample
{
    public static void Main()
    {
        // Prepare temporary file paths
        string inputDocxPath = Path.Combine(Path.GetTempPath(), "input.docx");
        string outputPdfPath = Path.Combine(Path.GetTempPath(), "output_pdfa.pdf");

        // -----------------------------------------------------------------
        // Create a sample DOCX document with some text (simulating source)
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Hello OCR - this document will be converted to PDF/A‑1a.");
        sourceDoc.Save(inputDocxPath, SaveFormat.Docx);

        // Verify the input DOCX was created
        if (!File.Exists(inputDocxPath))
            throw new InvalidOperationException("Failed to create the input DOCX file.");

        // ---------------------------------------------------------------
        // Load the DOCX and convert it to searchable PDF/A‑1a
        // ---------------------------------------------------------------
        Document doc = new Document(inputDocxPath);

        // Configure PDF/A‑1a compliance options
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA1a
        };

        // Perform the conversion
        doc.Save(outputPdfPath, pdfSaveOptions);

        // Validate that the searchable PDF/A‑1a file was created
        if (!File.Exists(outputPdfPath) || new FileInfo(outputPdfPath).Length == 0)
            throw new InvalidOperationException("The searchable PDF/A‑1a file was not created successfully.");

        // Clean up temporary files (optional)
        try { File.Delete(inputDocxPath); } catch { }

        // Indicate success (no interactive output required)
        Console.WriteLine("Searchable PDF/A‑1a conversion completed successfully.");
    }
}
