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
        // -----------------------------------------------------------------
        // Create a deterministic large image (>2 MB) using Aspose.Drawing.
        // Save as BMP to avoid compression that would keep the file small.
        // -----------------------------------------------------------------
        const string largeImagePath = "large.bmp";
        const int width = 3000;
        const int height = 3000;

        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height);
        Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap);
        graphics.Clear(Aspose.Drawing.Color.White);
        // Save as BMP – this format is uncompressed and will be larger than 2 MB.
        bitmap.Save(largeImagePath, ImageFormat.Bmp);
        graphics.Dispose();
        bitmap.Dispose();

        // Verify that the created image is larger than 2 MB.
        FileInfo fi = new FileInfo(largeImagePath);
        if (fi.Length <= 2 * 1024 * 1024)
            throw new Exception("Sample image is not larger than 2 MB.");

        // -----------------------------------------------------------------
        // Create a DOCX document and insert the large image.
        // -----------------------------------------------------------------
        const string inputDocPath = "input.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(largeImagePath);
        doc.Save(inputDocPath);

        // -----------------------------------------------------------------
        // Load the document for processing.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(inputDocPath);

        // Iterate over all Shape nodes that contain images.
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapeNodes)
        {
            if (!shape.HasImage)
                continue;

            // Get the original image bytes.
            byte[] originalBytes = shape.ImageData.ImageBytes;
            if (originalBytes.Length <= 2 * 1024 * 1024)
                continue; // Image already <= 2 MB.

            // Re‑encode the image as JPEG (compressed).
            using (MemoryStream inputStream = new MemoryStream(originalBytes))
            {
                using (Aspose.Drawing.Image img = Aspose.Drawing.Image.FromStream(inputStream))
                {
                    using (MemoryStream jpegStream = new MemoryStream())
                    {
                        img.Save(jpegStream, ImageFormat.Jpeg);
                        jpegStream.Position = 0; // Reset before reuse.

                        // Replace the shape's image with the compressed JPEG.
                        shape.ImageData.SetImage(jpegStream);
                    }
                }
            }
        }

        // -----------------------------------------------------------------
        // Save the modified document.
        // -----------------------------------------------------------------
        const string outputDocPath = "output.docx";
        loadedDoc.Save(outputDocPath);

        // Validate that the output file was created.
        if (!File.Exists(outputDocPath))
            throw new Exception("Failed to create the output document.");

        // Optional cleanup (commented out).
        // File.Delete(largeImagePath);
        // File.Delete(inputDocPath);
    }
}
