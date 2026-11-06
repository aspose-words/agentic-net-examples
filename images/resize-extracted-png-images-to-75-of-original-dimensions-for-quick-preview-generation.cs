using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a deterministic sample PNG image.
        const string sampleImagePath = "sample.png";
        const int originalWidth = 200;
        const int originalHeight = 200;

        using (Bitmap bitmap = new Bitmap(originalWidth, originalHeight))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                // Draw a simple red rectangle for visual distinction.
                g.FillRectangle(new SolidBrush(Color.Red), 20, 20, 160, 160);
            }
            bitmap.Save(sampleImagePath);
        }

        // Step 2: Create a Word document and insert the sample image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        const string docPath = "sample.docx";
        doc.Save(docPath);

        // Step 3: Load the document (already in memory) and extract PNG images.
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int previewIndex = 0;
        foreach (Shape shape in shapeNodes)
        {
            if (!shape.HasImage)
                continue;

            // Process only PNG images.
            if (shape.ImageData.ImageType != ImageType.Png)
                continue;

            // Save the original image to a memory stream.
            using (MemoryStream imageStream = new MemoryStream())
            {
                shape.ImageData.Save(imageStream);
                imageStream.Position = 0; // Reset before reading.

                // Load the image using Aspose.Drawing.
                using (Bitmap originalBitmap = new Bitmap(imageStream))
                {
                    // Calculate 75% dimensions.
                    int newWidth = (int)(originalBitmap.Width * 0.75);
                    int newHeight = (int)(originalBitmap.Height * 0.75);

                    // Create a resized bitmap.
                    using (Bitmap resizedBitmap = new Bitmap(newWidth, newHeight))
                    {
                        using (Graphics graphics = Graphics.FromImage(resizedBitmap))
                        {
                            graphics.DrawImage(originalBitmap, 0, 0, newWidth, newHeight);
                        }

                        // Save the preview image.
                        string previewPath = $"preview-{previewIndex}.png";
                        resizedBitmap.Save(previewPath);
                        previewIndex++;
                    }
                }
            }
        }

        // Validation: ensure at least one preview image was generated.
        if (previewIndex == 0)
            throw new Exception("No PNG images were extracted and resized.");

        // Optional: clean up sample files (comment out if inspection is needed).
        // File.Delete(sampleImagePath);
        // File.Delete(docPath);
    }
}
