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
        builder.Writeln("Sample document for PDF/A-2u conversion with image compression.");

        // Generate a simple image using Aspose.Drawing and insert it into the document.
        using (Bitmap bitmap = new Bitmap(100, 100))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.Blue);
            }

            using (MemoryStream imageStream = new MemoryStream())
            {
                bitmap.Save(imageStream, ImageFormat.Png);
                imageStream.Position = 0;
                builder.InsertImage(imageStream);
            }
        }

        // Configure PDF save options:
        // - PDF/A-2u compliance
        // - Automatic image compression
        // - JPEG quality for compressed images
        // - Optimize output to remove unused objects.
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA2u,
            ImageCompression = PdfImageCompression.Auto,
            JpegQuality = 80,
            OptimizeOutput = true
        };

        string outputPath = "output_pdfa2u.pdf";
        doc.Save(outputPath, saveOptions);

        // Verify that the file was created and is not empty.
        if (!File.Exists(outputPath) || new FileInfo(outputPath).Length == 0)
        {
            throw new InvalidOperationException("Failed to create the compressed PDF/A-2u file.");
        }
    }
}
