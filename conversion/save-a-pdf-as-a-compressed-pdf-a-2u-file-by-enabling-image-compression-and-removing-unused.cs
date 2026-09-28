using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample document in memory.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document to be saved as compressed PDF/A‑2u.");

        // Configure PDF save options:
        // - PDF/A‑2u compliance.
        // - JPEG image compression with quality 80.
        // - Optimize output (which also removes unused objects).
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA2u,
            ImageCompression = PdfImageCompression.Jpeg,
            JpegQuality = 80,
            OptimizeOutput = true
        };

        string outputPath = "compressed_pdfa2u.pdf";

        // Save the document as a compressed PDF/A‑2u file.
        doc.Save(outputPath, saveOptions);

        // Verify that the output file was created and is not empty.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Expected output file '{outputPath}' was not created.");
        }

        FileInfo fileInfo = new FileInfo(outputPath);
        if (fileInfo.Length == 0)
        {
            throw new InvalidOperationException($"Output file '{outputPath}' is empty.");
        }
    }
}
