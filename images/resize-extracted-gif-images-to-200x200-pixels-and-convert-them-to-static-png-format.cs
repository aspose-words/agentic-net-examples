using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Newtonsoft.Json; // Included as required package

public class Program
{
    public static void Main()
    {
        // Define deterministic file names
        const string gifPath = "sample.gif";
        const string docPath = "sample.docx";
        const string outputDir = "output";

        // Ensure output directory exists
        Directory.CreateDirectory(outputDir);

        // -------------------------------------------------
        // 1. Create a sample GIF image (static single‑frame)
        // -------------------------------------------------
        using (Bitmap bitmap = new Bitmap(100, 100))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(Color.Blue);
            // Draw a simple ellipse for visual content
            graphics.DrawEllipse(Pens.White, 10, 10, 80, 80);
            bitmap.Save(gifPath, ImageFormat.Gif);
        }

        // -------------------------------------------------
        // 2. Create a Word document and insert the GIF
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(gifPath);
        doc.Save(docPath);

        // -------------------------------------------------
        // 3. Load the document and extract GIF images
        // -------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;

        foreach (Shape shape in shapes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            if (shape.ImageData.ImageType != ImageType.Gif)
                continue;

            // Obtain the raw image bytes
            byte[] imageBytes = shape.ImageData.ToByteArray();

            // Load the GIF into Aspose.Drawing.Image
            using (MemoryStream ms = new MemoryStream(imageBytes))
            {
                ms.Position = 0;
                using (Image originalImage = Image.FromStream(ms))
                {
                    // Resize to 200x200 pixels
                    using (Bitmap resizedBitmap = new Bitmap(200, 200))
                    using (Graphics g = Graphics.FromImage(resizedBitmap))
                    {
                        g.DrawImage(originalImage, new Rectangle(0, 0, 200, 200));

                        // Save as static PNG
                        string outputPath = Path.Combine(outputDir, $"extracted_{imageIndex}.png");
                        resizedBitmap.Save(outputPath, ImageFormat.Png);

                        // Validation: ensure the file was created
                        if (!File.Exists(outputPath))
                            throw new InvalidOperationException($"Failed to create output file: {outputPath}");
                    }
                }
            }

            imageIndex++;
        }

        // Final validation: at least one PNG should have been produced
        if (imageIndex == 0)
            throw new InvalidOperationException("No GIF images were found and processed.");

        // Clean up temporary files (optional)
        // File.Delete(gifPath);
        // File.Delete(docPath);
    }
}
