using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Paths for temporary files
        string inputImagePath = "input.png";
        string documentPath = "doc_with_image.docx";
        string previewImagePath = "preview.png";

        // -------------------------------------------------
        // 1. Create a sample PNG image (200x200 white background)
        // -------------------------------------------------
        int originalWidth = 200;
        int originalHeight = 200;
        using (Bitmap bitmap = new Bitmap(originalWidth, originalHeight))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                // Draw a simple black rectangle for visual reference
                g.DrawRectangle(Pens.Black, 10, 10, originalWidth - 20, originalHeight - 20);
            }
            bitmap.Save(inputImagePath);
        }

        // -------------------------------------------------
        // 2. Create a Word document and insert the PNG image
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(inputImagePath);
        doc.Save(documentPath);

        // -------------------------------------------------
        // 3. Load the document and extract PNG images
        // -------------------------------------------------
        Document loadedDoc = new Document(documentPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        bool previewCreated = false;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Process only PNG images
            if (shape.ImageData.ImageType != ImageType.Png)
                continue;

            // Get image bytes from the shape
            byte[] imageBytes;
            using (MemoryStream ms = new MemoryStream())
            {
                shape.ImageData.Save(ms);
                imageBytes = ms.ToArray();
            }

            // Load the original image into a bitmap
            using (MemoryStream originalStream = new MemoryStream(imageBytes))
            using (Bitmap originalBitmap = new Bitmap(originalStream))
            {
                // Calculate 75% of original dimensions
                int newWidth = (int)(originalBitmap.Width * 0.75);
                int newHeight = (int)(originalBitmap.Height * 0.75);

                // Create a new bitmap with the reduced size
                using (Bitmap resizedBitmap = new Bitmap(newWidth, newHeight))
                {
                    using (Graphics g = Graphics.FromImage(resizedBitmap))
                    {
                        // Draw the original image scaled down into the new bitmap
                        g.DrawImage(originalBitmap, 0, 0, newWidth, newHeight);
                    }

                    // Save the resized preview image
                    resizedBitmap.Save(previewImagePath);
                    previewCreated = true;
                }
            }

            // Since the task mentions "extracted PNG images", we process each found PNG.
            // For this example there is only one, so we can break after processing.
            break;
        }

        // -------------------------------------------------
        // 4. Validate that the preview image was created
        // -------------------------------------------------
        if (!previewCreated || !File.Exists(previewImagePath))
            throw new InvalidOperationException("Preview image was not created.");

        // Cleanup: optional removal of temporary files can be added here if desired.
    }
}
