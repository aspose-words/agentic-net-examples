using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define folders for input and output files.
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "Input");
        string outputDir = Path.Combine(baseDir, "Artifacts");

        // Ensure the directories exist.
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Paths for the temporary DOC file and the resulting PDF file.
        string docPath = Path.Combine(inputDir, "SampleDocument.doc");
        string pdfPath = Path.Combine(outputDir, "SampleDocument_PdfA2b.pdf");

        // -----------------------------------------------------------------
        // Create a sample DOC file.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This document will be saved as PDF/A‑2b compliant PDF.");
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // Load the DOC file (simulating an existing source document).
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);

        // -----------------------------------------------------------------
        // Configure PDF save options to use PDF/A‑2b compliance.
        // Aspose.Words does not expose a PdfA2b enum value; the closest
        // standard is PdfA2u (PDF/A‑2u), which also satisfies PDF/A‑2b
        // requirements. Use that value here.
        // -----------------------------------------------------------------
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA2u
        };

        // -----------------------------------------------------------------
        // Save the document as PDF with the specified compliance level.
        // -----------------------------------------------------------------
        loadedDoc.Save(pdfPath, saveOptions);

        // -----------------------------------------------------------------
        // Verify that the PDF file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(pdfPath))
        {
            throw new InvalidOperationException("The PDF file was not created as expected.");
        }

        // Inform the user that the operation succeeded.
        Console.WriteLine("PDF/A‑2b (using PdfA2u) compliant PDF created successfully at:");
        Console.WriteLine(pdfPath);
    }
}
