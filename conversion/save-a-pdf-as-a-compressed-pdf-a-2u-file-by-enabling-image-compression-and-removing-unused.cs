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
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This PDF is saved as PDF/A‑2u with image compression and optimized output.");

        // Create a simple bitmap using Aspose.Drawing (no System.Drawing usage).
        using (Bitmap bitmap = new Bitmap(200, 200))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.CornflowerBlue);
                graphics.DrawEllipse(new Pen(Color.White, 5), 20, 20, 160, 160);
            }

            // Save the bitmap to a memory stream.
            using (MemoryStream imageStream = new MemoryStream())
            {
                bitmap.Save(imageStream, ImageFormat.Png);
                imageStream.Position = 0;

                // Insert the image into the document.
                builder.InsertImage(imageStream);
            }
        }

        // Configure PDF save options for PDF/A‑2u compliance, image compression, and output optimization.
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA2u,                 // PDF/A‑2u compliance.
            ImageCompression = PdfImageCompression.Jpeg,      // Compress images using JPEG.
            JpegQuality = 70,                                 // JPEG quality (0‑100, lower = higher compression).
            OptimizeOutput = true                             // Remove unused objects and optimize the PDF.
        };

        // Define the output file path.
        string outputPath = "CompressedPdfA2u.pdf";

        // Save the document as a compressed PDF/A‑2u file.
        doc.Save(outputPath, saveOptions);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The PDF/A‑2u file was not created.");

        // Optionally, you could check the file size to ensure compression took effect.
        long fileSize = new FileInfo(outputPath).Length;
        Console.WriteLine($"PDF/A‑2u file created successfully. Size: {fileSize} bytes.");
    }
}
