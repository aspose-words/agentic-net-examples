using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a deterministic sample PNG image.
        const string inputImagePath = "input.png";
        const int originalWidth = 200;
        const int originalHeight = 200;
        using (Bitmap bitmap = new Bitmap(originalWidth, originalHeight))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                // Draw a simple red rectangle.
                using (var pen = new Pen(Color.Red, 5))
                {
                    g.DrawRectangle(pen, 10, 10, originalWidth - 20, originalHeight - 20);
                }
            }
            bitmap.Save(inputImagePath);
        }

        // Create a document and insert the sample image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(inputImagePath);
        const string docPath = "original.docx";
        doc.Save(docPath);

        // Load the document and extract PNG images.
        Document loadedDoc = new Document(docPath);
        var shapes = loadedDoc.GetChildNodes(NodeType.Shape, true)
                              .Cast<Shape>()
                              .Where(s => s.HasImage && s.ImageData.ImageType == ImageType.Png)
                              .ToList();

        int previewCount = 0;
        foreach (var shape in shapes)
        {
            // Save the original image to a memory stream.
            using (MemoryStream ms = new MemoryStream())
            {
                shape.ImageData.Save(ms);
                ms.Position = 0; // Reset before reading.

                // Load the image with Aspose.Drawing.
                using (Bitmap originalBitmap = new Bitmap(ms))
                {
                    int newWidth = originalBitmap.Width / 2;
                    int newHeight = originalBitmap.Height / 2;

                    // Create a resized bitmap.
                    using (Bitmap resizedBitmap = new Bitmap(newWidth, newHeight))
                    {
                        using (Graphics graphics = Graphics.FromImage(resizedBitmap))
                        {
                            graphics.Clear(Color.White);
                            graphics.DrawImage(originalBitmap, new Rectangle(0, 0, newWidth, newHeight));
                        }

                        // Save the preview image.
                        string previewPath = $"preview_{previewCount}.png";
                        resizedBitmap.Save(previewPath);
                        previewCount++;
                    }
                }
            }
        }

        // Validate that at least one preview image was generated.
        if (previewCount == 0)
            throw new InvalidOperationException("No PNG images were extracted and resized.");

        // Optional: clean up intermediate files (commented out if you want to keep them).
        // File.Delete(inputImagePath);
        // File.Delete(docPath);
    }
}
