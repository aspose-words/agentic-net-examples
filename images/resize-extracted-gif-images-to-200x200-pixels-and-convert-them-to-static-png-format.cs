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
        // Directories for artifacts
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // 1. Create a sample GIF image (static for simplicity)
        string gifPath = Path.Combine(artifactsDir, "sample.gif");
        using (Bitmap bmp = new Bitmap(300, 300))
        {
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.Clear(Color.White);
                // Draw a simple rectangle to have visible content
                g.FillRectangle(Brushes.Blue, 50, 50, 200, 200);
            }
            bmp.Save(gifPath, ImageFormat.Gif);
        }

        // 2. Create a Word document and insert the GIF image
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(gifPath);
        string docPath = Path.Combine(artifactsDir, "DocumentWithGif.docx");
        doc.Save(docPath);

        // 3. Load the document and extract GIF images
        Document loadedDoc = new Document(docPath);
        NodeCollection shapes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        foreach (Shape shape in shapes)
        {
            if (!shape.HasImage) continue;
            if (shape.ImageData.ImageType != ImageType.Gif) continue;

            // 4. Get image bytes and load into Aspose.Drawing.Bitmap
            byte[] gifBytes = shape.ImageData.ToByteArray();
            using (MemoryStream ms = new MemoryStream(gifBytes))
            {
                ms.Position = 0;
                using (Bitmap originalBmp = new Bitmap(ms))
                {
                    // 5. Resize to 200x200 pixels
                    using (Bitmap resizedBmp = new Bitmap(200, 200))
                    {
                        using (Graphics g = Graphics.FromImage(resizedBmp))
                        {
                            g.Clear(Color.Transparent);
                            g.DrawImage(originalBmp, new Rectangle(0, 0, 200, 200));
                        }

                        // 6. Save as static PNG
                        string pngPath = Path.Combine(artifactsDir, $"ExtractedImage_{imageIndex}.png");
                        resizedBmp.Save(pngPath, ImageFormat.Png);

                        // Validate that the PNG was created
                        if (!File.Exists(pngPath))
                            throw new InvalidOperationException($"Failed to create PNG file: {pngPath}");
                    }
                }
            }

            imageIndex++;
        }

        // Ensure at least one image was processed
        if (imageIndex == 0)
            throw new InvalidOperationException("No GIF images were found in the document.");

        // Optional: clean up intermediate files (commented out if you want to inspect them)
        // File.Delete(gifPath);
        // File.Delete(docPath);
    }
}
