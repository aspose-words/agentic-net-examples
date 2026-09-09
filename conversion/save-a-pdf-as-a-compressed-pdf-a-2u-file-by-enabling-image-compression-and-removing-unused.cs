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
        // Define file paths.
        const string outputPdfPath = "CompressedPdfA2u.pdf";

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample content for PDF/A‑2u conversion with image compression.");

        // Create a simple bitmap using Aspose.Drawing (no System.Drawing usage).
        using (Bitmap bitmap = new Bitmap(200, 200))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.Blue);
            }

            // Save the bitmap to a memory stream in PNG format.
            using (MemoryStream imageStream = new MemoryStream())
            {
                bitmap.Save(imageStream, ImageFormat.Png);
                imageStream.Position = 0; // Reset position before reading.

                // Insert the image into the document.
                builder.InsertImage(imageStream);
            }
        }

        // Configure PDF save options for PDF/A‑2u compliance,
        // enable JPEG image compression, set quality, and remove unused objects.
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA2u,
            ImageCompression = PdfImageCompression.Jpeg,
            JpegQuality = 75,
            OptimizeOutput = true // Removes redundant objects.
        };

        // Save the document as a compressed PDF/A‑2u file.
        doc.Save(outputPdfPath, saveOptions);

        // Validate that the output file was created and is not empty.
        if (!File.Exists(outputPdfPath) || new FileInfo(outputPdfPath).Length == 0)
        {
            throw new InvalidOperationException("The PDF/A‑2u file was not created or is empty.");
        }

        // Optionally, inform that the process completed successfully.
        Console.WriteLine($"PDF/A‑2u file saved successfully: {outputPdfPath}");
    }
}
