using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Create a deterministic sample image (PNG) using Aspose.Drawing.
        const string sampleImagePath = "sample.png";
        const int imgWidth = 100;
        const int imgHeight = 100;

        using (var bitmap = new Bitmap(imgWidth, imgHeight))
        using (var graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(Color.White);
            // Draw a simple rectangle.
            graphics.DrawRectangle(Pens.Black, 10, 10, imgWidth - 20, imgHeight - 20);
            bitmap.Save(sampleImagePath, ImageFormat.Png);
        }

        // Create a new Word document and insert the sample image.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);

        // Save the document (optional, just to have a file on disk).
        const string docPath = "sample.docx";
        doc.Save(docPath);

        // Extract the image from the document.
        Shape imageShape = (Shape)doc.GetChild(NodeType.Shape, 0, true);
        if (imageShape == null || !imageShape.HasImage)
            throw new InvalidOperationException("No image found in the document.");

        // Save the extracted image to a temporary memory stream (original format).
        using (var originalImageStream = new MemoryStream())
        {
            imageShape.ImageData.Save(originalImageStream);
            originalImageStream.Position = 0;

            // Load the image into Aspose.Drawing.Bitmap to convert it to BMP.
            using (var bitmap = new Bitmap(originalImageStream))
            using (var bmpStream = new MemoryStream())
            {
                // Save as BMP into the memory stream.
                bitmap.Save(bmpStream, ImageFormat.Bmp);
                bmpStream.Position = 0;

                // Pass the BMP memory stream to the API (demo method).
                ProcessImageStream(bmpStream);
            }
        }

        // Clean up temporary files.
        if (File.Exists(sampleImagePath)) File.Delete(sampleImagePath);
        if (File.Exists(docPath)) File.Delete(docPath);
    }

    // Dummy API method that receives a BMP image stream.
    private static void ProcessImageStream(Stream bmpStream)
    {
        // Validate that the stream contains data.
        if (bmpStream == null || !bmpStream.CanRead)
            throw new ArgumentException("Invalid image stream.");

        // For demonstration, write the stream to a file.
        const string outputPath = "extracted.bmp";
        using (var fileStream = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
        {
            bmpStream.CopyTo(fileStream);
        }

        // Verify that the file was created.
        if (!File.Exists(outputPath) || new FileInfo(outputPath).Length == 0)
            throw new InvalidOperationException("Failed to write the BMP image to disk.");
    }
}
