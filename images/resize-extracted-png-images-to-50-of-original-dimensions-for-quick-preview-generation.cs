using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Aspose.Drawing.Drawing2D; // For InterpolationMode

public class Program
{
    public static void Main()
    {
        // Deterministic file names
        const string inputImagePath = "input.png";
        const string docPath = "original.docx";
        const string previewPrefix = "preview_";

        // 1. Create a sample PNG image (200x200) and save it as input.png
        const int sampleWidth = 200;
        const int sampleHeight = 200;
        using (Bitmap bitmap = new Bitmap(sampleWidth, sampleHeight))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(Color.White);
            using (Pen pen = new Pen(Color.Red, 5))
            {
                graphics.DrawRectangle(pen, 10, 10, sampleWidth - 20, sampleHeight - 20);
            }
            bitmap.Save(inputImagePath);
        }

        // 2. Insert the image into a new Word document
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(inputImagePath);
        doc.Save(docPath);

        // 3. Load the document and process each PNG image
        Document loadedDoc = new Document(docPath);
        NodeCollection shapes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int previewIndex = 0;

        foreach (Shape shape in shapes)
        {
            if (!shape.HasImage)
                continue;

            // Process only PNG images
            if (shape.ImageData.ImageType != ImageType.Png)
                continue;

            // Extract the image bytes into a memory stream
            using (MemoryStream imageStream = new MemoryStream())
            {
                shape.ImageData.Save(imageStream);
                imageStream.Position = 0;

                // Load the image using Aspose.Drawing
                using (Bitmap originalBitmap = new Bitmap(imageStream))
                {
                    // Calculate 50% dimensions
                    int newWidth = originalBitmap.Width / 2;
                    int newHeight = originalBitmap.Height / 2;

                    // Create a new bitmap for the resized preview
                    using (Bitmap previewBitmap = new Bitmap(newWidth, newHeight))
                    using (Graphics g = Graphics.FromImage(previewBitmap))
                    {
                        // High-quality scaling
                        g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                        g.DrawImage(originalBitmap, new Rectangle(0, 0, newWidth, newHeight));

                        // Save the preview image
                        string previewPath = $"{previewPrefix}{previewIndex}.png";
                        previewBitmap.Save(previewPath);
                        previewIndex++;
                    }
                }
            }
        }

        // Validation: ensure at least one preview was generated
        if (previewIndex == 0)
            throw new InvalidOperationException("No PNG images were found to generate previews.");

        // Optional cleanup (commented out to keep artifacts)
        // File.Delete(inputImagePath);
        // File.Delete(docPath);
    }
}
