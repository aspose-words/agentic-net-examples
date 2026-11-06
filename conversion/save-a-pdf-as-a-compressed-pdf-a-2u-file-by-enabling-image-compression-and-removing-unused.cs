using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document saved as compressed PDF/A‑2u.");

        // Configure PDF save options for PDF/A‑2u compliance with image compression and output optimization.
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA2u,
            ImageCompression = PdfImageCompression.Jpeg,
            JpegQuality = 80,
            OptimizeOutput = true
        };

        // Define output file path.
        string outputPath = "compressed_pdfa2u.pdf";

        // Save the document as a compressed PDF/A‑2u file.
        doc.Save(outputPath, saveOptions);

        // Verify that the output file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Expected output file '{outputPath}' was not created.");
        }

        // Verify that the file size is greater than zero.
        FileInfo info = new FileInfo(outputPath);
        if (info.Length == 0)
        {
            throw new InvalidOperationException("The generated PDF file is empty.");
        }
    }
}
