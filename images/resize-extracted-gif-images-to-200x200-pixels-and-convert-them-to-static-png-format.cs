using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Deterministic file names for the sample GIF and the Word document.
        const string gifPath = "sample.gif";
        const string docPath = "document_with_gif.docx";

        // 1. Create a sample 300x300 GIF image using Aspose.Drawing and save it.
        using (Aspose.Drawing.Bitmap bmp = new Aspose.Drawing.Bitmap(300, 300))
        {
            using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bmp))
            {
                g.Clear(Aspose.Drawing.Color.White);
                g.DrawRectangle(Aspose.Drawing.Pens.Black, 50, 50, 200, 200);
                g.DrawString("GIF", new Aspose.Drawing.Font("Arial", 48), Aspose.Drawing.Brushes.Black,
                    new Aspose.Drawing.PointF(80, 120));
            }

            bmp.Save(gifPath, ImageFormat.Gif);
        }

        // 2. Create a new Word document and insert the GIF image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(gifPath);
        doc.Save(docPath);

        // 3. Load the document and extract GIF images.
        Document loadedDoc = new Document(docPath);
        NodeCollection shapes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;

        foreach (Shape shape in shapes.OfType<Shape>())
        {
            if (shape.HasImage && shape.ImageData.ImageType == ImageType.Gif)
            {
                // 4. Obtain the image bytes from the shape.
                byte[] imageBytes;
                using (MemoryStream ms = new MemoryStream())
                {
                    shape.ImageData.Save(ms);
                    ms.Position = 0;
                    imageBytes = ms.ToArray();
                }

                // 5. Load the GIF into a bitmap, resize to 200x200, and save as PNG.
                using (MemoryStream msInput = new MemoryStream(imageBytes))
                using (Aspose.Drawing.Bitmap original = new Aspose.Drawing.Bitmap(msInput))
                {
                    using (Aspose.Drawing.Bitmap resized = new Aspose.Drawing.Bitmap(200, 200))
                    {
                        using (Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(resized))
                        {
                            graphics.Clear(Aspose.Drawing.Color.White);
                            graphics.DrawImage(original, new Aspose.Drawing.Rectangle(0, 0, 200, 200));
                        }

                        string outputPath = $"extracted_{extractedCount}.png";
                        resized.Save(outputPath, ImageFormat.Png);
                        extractedCount++;
                    }
                }
            }
        }

        // 6. Validate that at least one PNG was created.
        if (extractedCount == 0)
            throw new InvalidOperationException("No GIF images were extracted and converted.");

        // Optional cleanup (commented out).
        // File.Delete(gifPath);
        // File.Delete(docPath);
    }
}
