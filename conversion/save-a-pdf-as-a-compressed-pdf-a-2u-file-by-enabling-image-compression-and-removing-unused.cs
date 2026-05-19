using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Drawing2D;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Create a new blank Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This document will be saved as a compressed PDF/A‑2u file.");

        // Generate a simple image using Aspose.Drawing.
        using (Bitmap bitmap = new Bitmap(200, 200))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.LightBlue);
                using (Pen pen = new Pen(Color.DarkBlue, 5))
                {
                    graphics.DrawEllipse(pen, 20, 20, 160, 160);
                }
            }

            using (MemoryStream imageStream = new MemoryStream())
            {
                // Save the bitmap to a memory stream as PNG.
                bitmap.Save(imageStream, ImageFormat.Png);
                imageStream.Position = 0;

                // Insert the image into the document.
                builder.InsertImage(imageStream);
            }
        }

        // Configure PDF save options:
        // - PDF/A‑2u compliance.
        // - JPEG image compression with quality 80.
        // - Optimize output to remove unused objects.
        // - Apply Flate text compression.
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA2u,
            ImageCompression = PdfImageCompression.Jpeg,
            JpegQuality = 80,
            OptimizeOutput = true,
            TextCompression = PdfTextCompression.Flate
        };

        string outputPath = "CompressedPdfA2u.pdf";

        // Save the document as a PDF/A‑2u file with the specified options.
        doc.Save(outputPath, saveOptions);

        // Verify that the output file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The PDF/A‑2u file was not created.");

        // Output the size of the generated file.
        FileInfo fileInfo = new FileInfo(outputPath);
        Console.WriteLine($"PDF/A‑2u file created: {outputPath} ({fileInfo.Length} bytes)");
    }
}
