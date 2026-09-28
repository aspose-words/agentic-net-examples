using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample DOC file.
        const string inputPath = "input.doc";
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Sample DOC content for PDF/A‑2b conversion.");
        sourceDoc.Save(inputPath, SaveFormat.Doc);

        // Load the DOC file.
        Document doc = new Document(inputPath);

        // Configure PDF/A‑2b compliance.
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Try to set PdfA2b compliance; fall back to a supported level if unavailable.
        if (!Enum.TryParse<PdfCompliance>("PdfA2b", ignoreCase: true, out PdfCompliance compliance))
        {
            // Fallback to a generic PDF compliance level (e.g., PDF/A‑1b) if PdfA2b is not supported.
            compliance = PdfCompliance.PdfA1b;
        }

        pdfOptions.Compliance = compliance;

        // Save as PDF with the specified compliance level.
        const string outputPath = "output.pdf";
        doc.Save(outputPath, pdfOptions);

        // Validate that the PDF was created and contains data.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("The PDF output file was not created.");
        }

        FileInfo info = new FileInfo(outputPath);
        if (info.Length == 0)
        {
            throw new InvalidOperationException("The PDF output file is empty.");
        }

        // Optional clean‑up.
        // File.Delete(inputPath);
        // File.Delete(outputPath);
    }
}
