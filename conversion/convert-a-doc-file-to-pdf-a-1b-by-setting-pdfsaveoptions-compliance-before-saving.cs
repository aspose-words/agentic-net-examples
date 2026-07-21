using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample DOC file.
        const string inputPath = "sample.doc";
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is a sample document for PDF/A‑1b conversion.");
        sourceDoc.Save(inputPath, SaveFormat.Doc);

        // Load the DOC file.
        Document doc = new Document(inputPath);

        // Configure PDF save options for PDF/A‑1b compliance.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA1b
        };

        // Save the document as PDF/A‑1b.
        const string outputPath = "output.pdf";
        doc.Save(outputPath, pdfOptions);

        // Verify that the output file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The PDF/A‑1b file was not created.");

        // Optional: clean up the sample input file.
        // File.Delete(inputPath);
    }
}
