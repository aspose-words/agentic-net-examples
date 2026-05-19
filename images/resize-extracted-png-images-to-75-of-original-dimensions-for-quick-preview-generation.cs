using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Loading;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Drawing2D; // For InterpolationMode

public class Program
{
    public static void Main()
    {
        // Define file names
        const string inputImagePath = "input.png";
        const string docPath = "document.docx";

        // -------------------------------------------------
        // 1. Create a sample PNG image (200x200) using Aspose.Drawing
        // -------------------------------------------------
        int originalWidth = 200;
        int originalHeight = 200;
        using (Bitmap bitmap = new Bitmap(originalWidth, originalHeight))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                // Fill background with white
                g.Clear(Color.White);
                // Draw a simple red rectangle
                using (var pen = new Pen(Color.Red, 5))
                {
                    g.DrawRectangle(pen, 10, 10, originalWidth - 20, originalHeight - 20);
                }
            }
            // Save the created image
            bitmap.Save(inputImagePath);
        }

        // -------------------------------------------------
        // 2. Create a Word document and insert the PNG image
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(inputImagePath);
        doc.Save(docPath);

        // -------------------------------------------------
        // 3. Load the document and extract images
        // -------------------------------------------------
        Document loadedDoc = new Document(docPath);
        var shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true)
                                  .Cast<Shape>()
                                  .Where(s => s.HasImage)
                                  .ToList();

        if (!shapeNodes.Any())
            throw new InvalidOperationException("No images were found in the document.");

        int imageIndex = 0;
        foreach (var shape in shapeNodes)
        {
            // Save the original extracted image
            string extractedPath = $"extracted_{imageIndex}.png";
            shape.ImageData.Save(extractedPath);

            // -------------------------------------------------
            // 4. Resize the extracted PNG to 75% of original dimensions
            // -------------------------------------------------
            using (Bitmap originalBitmap = new Bitmap(extractedPath))
            {
                int newWidth = (int)(originalBitmap.Width * 0.75);
                int newHeight = (int)(originalBitmap.Height * 0.75);

                using (Bitmap resizedBitmap = new Bitmap(newWidth, newHeight))
                {
                    using (Graphics g = Graphics.FromImage(resizedBitmap))
                    {
                        // High quality scaling
                        g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                        g.DrawImage(originalBitmap, 0, 0, newWidth, newHeight);
                    }

                    // Save the preview image
                    string previewPath = $"preview_{imageIndex}.png";
                    resizedBitmap.Save(previewPath);

                    // Validate that the preview file was created
                    if (!File.Exists(previewPath) || new FileInfo(previewPath).Length == 0)
                        throw new InvalidOperationException($"Failed to create preview image: {previewPath}");
                }
            }

            imageIndex++;
        }

        // -------------------------------------------------
        // 5. Clean up (optional) - all using statements already disposed resources
        // -------------------------------------------------
        Console.WriteLine("Image preview generation completed successfully.");
    }
}
