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
        // Paths for temporary files
        const string largeImagePath = "large.png";
        const string inputDocPath = "input.docx";
        const string outputDocPath = "output.docx";

        // -------------------------------------------------
        // Step 1: Create a large sample image (>2 MB)
        // -------------------------------------------------
        int width = 3000;
        int height = 3000;
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Aspose.Drawing.Color.White);
            }
            // Save as PNG to ensure large file size
            bitmap.Save(largeImagePath, ImageFormat.Png);
        }

        // Verify that the image file was created
        if (!File.Exists(largeImagePath))
            throw new FileNotFoundException("Failed to create the sample image.", largeImagePath);

        // -------------------------------------------------
        // Step 2: Create a DOCX document and insert the image
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(largeImagePath);
        doc.Save(inputDocPath);

        // Verify that the input document was created
        if (!File.Exists(inputDocPath))
            throw new FileNotFoundException("Failed to create the input document.", inputDocPath);

        // -------------------------------------------------
        // Step 3: Load the document and replace large images
        // -------------------------------------------------
        Document loadedDoc = new Document(inputDocPath);
        NodeCollection shapes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        foreach (Shape shape in shapes)
        {
            if (!shape.HasImage)
                continue;

            // Get the original image bytes
            byte[] originalBytes = shape.ImageData.ImageBytes;

            // Check if image size exceeds 2 MB
            const long twoMegabytes = 2L * 1024L * 1024L;
            if (originalBytes.Length <= twoMegabytes)
                continue;

            // Load the image into a bitmap
            using (MemoryStream originalStream = new MemoryStream(originalBytes))
            {
                using (Bitmap bitmap = new Bitmap(originalStream))
                {
                    // Re-encode the bitmap as JPEG (compressed)
                    using (MemoryStream jpegStream = new MemoryStream())
                    {
                        bitmap.Save(jpegStream, ImageFormat.Jpeg);
                        jpegStream.Position = 0; // Reset for reading

                        // Replace the shape's image with the compressed JPEG
                        shape.ImageData.SetImage(jpegStream);
                    }
                }
            }
        }

        // -------------------------------------------------
        // Step 4: Save the modified document
        // -------------------------------------------------
        loadedDoc.Save(outputDocPath);

        // Validate that the output document exists
        if (!File.Exists(outputDocPath))
            throw new FileNotFoundException("Failed to save the output document.", outputDocPath);
    }
}
