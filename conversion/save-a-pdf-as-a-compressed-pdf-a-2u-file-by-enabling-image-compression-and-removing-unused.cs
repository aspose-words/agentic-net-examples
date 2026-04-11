using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Prepare folders.
        string baseDir = Directory.GetCurrentDirectory();
        string artifactsDir = Path.Combine(baseDir, "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string imagePath = Path.Combine(artifactsDir, "sample.png");

        // Create a simple bitmap using Aspose.Drawing and save it.
        using (Bitmap bitmap = new Bitmap(200, 200, PixelFormat.Format24bppRgb))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.Blue);
            }
            bitmap.Save(imagePath, ImageFormat.Png);
        }

        // Build a sample Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample document for PDF/A‑2u conversion with compression.");
        builder.InsertImage(imagePath);
        builder.Writeln("End of document.");

        // Configure PDF save options for PDF/A‑2u, image compression, and unused‑object removal.
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA2u,                 // PDF/A‑2u compliance.
            ImageCompression = PdfImageCompression.Jpeg,      // Compress all images as JPEG.
            JpegQuality = 70,                                 // Adjust JPEG quality for stronger compression.
            OptimizeOutput = true                             // Remove unused objects from the PDF.
        };

        // Save the document.
        string pdfPath = Path.Combine(artifactsDir, "CompressedPdfA2u.pdf");
        doc.Save(pdfPath, saveOptions);

        // Verify that the output file was created and is not empty.
        if (!File.Exists(pdfPath) || new FileInfo(pdfPath).Length == 0)
        {
            throw new InvalidOperationException("Failed to create the compressed PDF/A‑2u file.");
        }

        // Optional: indicate success (no console interaction required by the task).
    }
}
