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
        // Create a deterministic sample PNG image.
        const string inputImagePath = "sample.png";
        const int originalWidth = 200;
        const int originalHeight = 200;

        using (Bitmap bitmap = new Bitmap(originalWidth, originalHeight))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                // Fill background with white.
                g.Clear(Color.White);
                // Draw a simple red rectangle.
                using (SolidBrush brush = new SolidBrush(Color.Red))
                {
                    g.FillRectangle(brush, 20, 20, originalWidth - 40, originalHeight - 40);
                }
            }
            bitmap.Save(inputImagePath, ImageFormat.Png);
        }

        // Create a Word document and insert the sample image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(inputImagePath);
        const string docPath = "DocumentWithImage.docx";
        doc.Save(docPath);

        // Reload the document (optional, demonstrates load step).
        Document loadedDoc = new Document(docPath);

        // Extract PNG images, resize them to 50%, and save previews.
        int previewIndex = 0;
        foreach (Shape shape in loadedDoc.GetChildNodes(NodeType.Shape, true))
        {
            if (!shape.HasImage)
                continue;

            // Only process PNG images.
            if (shape.ImageData.ImageType != ImageType.Png)
                continue;

            // Save the original image to a memory stream.
            using (MemoryStream originalStream = new MemoryStream())
            {
                shape.ImageData.Save(originalStream);
                originalStream.Position = 0; // Reset before reading.

                // Load the image with Aspose.Drawing.
                using (Bitmap originalBitmap = new Bitmap(originalStream))
                {
                    int newWidth = originalBitmap.Width / 2;
                    int newHeight = originalBitmap.Height / 2;

                    // Create a new bitmap for the resized preview.
                    using (Bitmap resizedBitmap = new Bitmap(newWidth, newHeight))
                    {
                        using (Graphics graphics = Graphics.FromImage(resizedBitmap))
                        {
                            // Draw the original image scaled down.
                            graphics.DrawImage(originalBitmap, 0, 0, newWidth, newHeight);
                        }

                        // Save the resized preview.
                        string previewPath = $"preview-{previewIndex}.png";
                        resizedBitmap.Save(previewPath, ImageFormat.Png);
                        previewIndex++;
                    }
                }
            }
        }

        // Validate that at least one preview image was created.
        if (previewIndex == 0)
            throw new InvalidOperationException("No PNG images were extracted and resized.");

        // Optional cleanup: ensure the sample files exist.
        if (!File.Exists(inputImagePath) || !File.Exists(docPath))
            throw new FileNotFoundException("Required input files were not created.");

        // The program finishes automatically.
    }
}
