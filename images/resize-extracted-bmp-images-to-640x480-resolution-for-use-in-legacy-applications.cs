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
        // ------------------------------------------------------------
        // 1. Create a deterministic sample BMP image (200x200).
        // ------------------------------------------------------------
        const string sampleBmpPath = "sample.bmp";
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(200, 200);
        Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap);
        g.Clear(Aspose.Drawing.Color.White);
        // Draw a simple blue rectangle.
        using (Aspose.Drawing.SolidBrush brush = new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.Blue))
        {
            g.FillRectangle(brush, 20, 20, 160, 160);
        }
        // Save the bitmap to a file so it can be inserted into the document.
        bitmap.Save(sampleBmpPath);
        // Clean up drawing objects.
        g.Dispose();
        bitmap.Dispose();

        // ------------------------------------------------------------
        // 2. Create a Word document and insert the BMP image.
        // ------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleBmpPath);
        const string docPath = "sample.docx";
        doc.Save(docPath);

        // ------------------------------------------------------------
        // 3. Reload the document to simulate an extraction scenario.
        // ------------------------------------------------------------
        Document loadedDoc = new Document(docPath);

        // ------------------------------------------------------------
        // 4. Extract each image, resize it to 640x480, and save the result.
        // ------------------------------------------------------------
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes)
        {
            if (!shape.HasImage)
                continue;

            // Save the original image data to a memory stream.
            using (MemoryStream originalStream = new MemoryStream())
            {
                // Aspose.Words ImageData.Save overload takes only the stream.
                shape.ImageData.Save(originalStream);
                originalStream.Position = 0; // Reset before reading.

                // Load the image into an Aspose.Drawing.Bitmap.
                using (Aspose.Drawing.Bitmap originalBitmap = new Aspose.Drawing.Bitmap(originalStream))
                {
                    // Create a new bitmap with the target size.
                    using (Aspose.Drawing.Bitmap resizedBitmap = new Aspose.Drawing.Bitmap(640, 480))
                    {
                        using (Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(resizedBitmap))
                        {
                            // Draw the original image stretched to the new dimensions.
                            graphics.DrawImage(originalBitmap, new Rectangle(0, 0, 640, 480));
                        }

                        string resizedPath = $"resized-{imageIndex}.bmp";
                        resizedBitmap.Save(resizedPath);
                        if (!File.Exists(resizedPath))
                            throw new Exception($"Resized image was not created: {resizedPath}");
                    }
                }
            }

            imageIndex++;
        }

        // ------------------------------------------------------------
        // 5. Validate that at least one image was processed.
        // ------------------------------------------------------------
        if (imageIndex == 0)
            throw new Exception("No images were extracted from the document.");
    }
}
