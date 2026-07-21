using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Loading;
using Aspose.Words.Saving;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // 1. Create a sample PNG image (200x200) and save it as input.png.
        string inputImagePath = Path.Combine(artifactsDir, "input.png");
        using (Bitmap bitmap = new Bitmap(200, 200))
        using (Graphics g = Graphics.FromImage(bitmap))
        {
            g.Clear(Color.White);
            // Draw a simple red rectangle for visual distinction.
            using (var pen = new Pen(Color.Red, 5))
            {
                g.DrawRectangle(pen, 10, 10, 180, 180);
            }
            bitmap.Save(inputImagePath);
        }

        // 2. Create a new Word document and insert the sample image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(inputImagePath);
        string docPath = Path.Combine(artifactsDir, "original.docx");
        doc.Save(docPath);

        // 3. Load the document and extract PNG images.
        LoadOptions loadOptions = new LoadOptions();
        Document loadedDoc = new Document(docPath, loadOptions);
        var shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true)
                                  .OfType<Shape>()
                                  .Where(s => s.HasImage && s.ImageData.ImageType == ImageType.Png)
                                  .ToList();

        int previewIndex = 0;
        foreach (var shape in shapeNodes)
        {
            // Save the image data to a memory stream.
            using (MemoryStream imageStream = new MemoryStream())
            {
                shape.ImageData.Save(imageStream);
                imageStream.Position = 0;

                // Load the original bitmap.
                using (Bitmap originalBitmap = new Bitmap(imageStream))
                {
                    // Calculate new dimensions (50% of original).
                    int newWidth = originalBitmap.Width / 2;
                    int newHeight = originalBitmap.Height / 2;

                    // Create a resized bitmap.
                    using (Bitmap resizedBitmap = new Bitmap(newWidth, newHeight))
                    using (Graphics graphics = Graphics.FromImage(resizedBitmap))
                    {
                        graphics.DrawImage(originalBitmap, 0, 0, newWidth, newHeight);
                        string previewPath = Path.Combine(artifactsDir, $"preview_{previewIndex}.png");
                        resizedBitmap.Save(previewPath);
                        previewIndex++;
                    }
                }
            }
        }

        // Validate that at least one preview image was created.
        if (previewIndex == 0)
            throw new Exception("No PNG images were extracted and resized.");

        // Optional: indicate completion (no interactive prompts).
        Console.WriteLine($"Resized {previewIndex} preview image(s) saved to '{artifactsDir}'.");
    }
}
