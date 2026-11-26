using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Aspose.Drawing.Drawing2D;

public class Program
{
    public static void Main()
    {
        // -------------------------------------------------
        // 1. Create a deterministic sample PNG image.
        // -------------------------------------------------
        const string sampleImagePath = "input.png";
        using (var bitmap = new Bitmap(200, 200))
        {
            using (var graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.White);
                using (var pen = new Pen(Color.Blue, 3))
                {
                    graphics.DrawEllipse(pen, 20, 20, 160, 160);
                }
            }
            bitmap.Save(sampleImagePath, ImageFormat.Png);
        }

        // -------------------------------------------------
        // 2. Insert the sample PNG image twice into a Word document.
        // -------------------------------------------------
        const string docPath = "sample.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        builder.Writeln(); // separate images
        builder.InsertImage(sampleImagePath);
        doc.Save(docPath);

        // -------------------------------------------------
        // 3. Reload the document and extract PNG images.
        // -------------------------------------------------
        var loadedDoc = new Document(docPath);
        var shapes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;

        for (int i = 0; i < shapes.Count; i++)
        {
            var shape = (Shape)shapes[i];
            if (!shape.HasImage)
                continue;

            if (shape.ImageData.ImageType != ImageType.Png)
                continue;

            // Extract the image to a memory stream.
            using (var ms = new MemoryStream())
            {
                shape.ImageData.Save(ms);
                ms.Position = 0;

                // Load the extracted image.
                using (var original = new Bitmap(ms))
                {
                    // Create a new non‑indexed bitmap (32‑bpp ARGB) to allow drawing.
                    using (var target = new Bitmap(original.Width, original.Height, PixelFormat.Format32bppArgb))
                    {
                        // Copy original image onto the new bitmap.
                        using (var g = Graphics.FromImage(target))
                        {
                            g.DrawImage(original, 0, 0, original.Width, original.Height);

                            // Draw a 5‑pixel red border.
                            using (var pen = new Pen(Color.Red, 5))
                            {
                                // Rectangle inside the image bounds.
                                g.DrawRectangle(pen, 0, 0, target.Width - 1, target.Height - 1);
                            }
                        }

                        // Save the modified image.
                        string outPath = $"extracted_{extractedCount + 1}.png";
                        target.Save(outPath, ImageFormat.Png);
                        extractedCount++;
                    }
                }
            }
        }

        // -------------------------------------------------
        // 4. Validate that at least one image was processed.
        // -------------------------------------------------
        if (extractedCount == 0)
            throw new InvalidOperationException("No PNG images were extracted and processed.");
    }
}
