using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Prepare folders.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        string inputDir = Path.Combine(artifactsDir, "InputImages");
        string outputDir = Path.Combine(artifactsDir, "OutputImages");
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a sample GIF image (static for simplicity) and save it.
        // -----------------------------------------------------------------
        string gifPath = Path.Combine(inputDir, "sample.gif");
        using (Bitmap bmp = new Bitmap(200, 200))
        using (Graphics g = Graphics.FromImage(bmp))
        {
            g.Clear(Aspose.Drawing.Color.LightBlue);
            bmp.Save(gifPath, ImageFormat.Gif);
        }

        // --------------------------------------------------------------
        // 2. Create a Word document and insert the GIF image into it.
        // --------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(gifPath);
        string docPath = Path.Combine(artifactsDir, "DocumentWithGif.docx");
        doc.Save(docPath);

        // --------------------------------------------------------------
        // 3. Load the document and extract all GIF images.
        // --------------------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage) continue;
            if (shape.ImageData.ImageType != ImageType.Gif) continue;

            // Save the GIF image to a memory stream.
            using (MemoryStream gifStream = new MemoryStream())
            {
                shape.ImageData.Save(gifStream);
                gifStream.Position = 0;

                // Load the GIF using Aspose.Drawing.Image.
                using (Image gifImage = Image.FromStream(gifStream))
                {
                    // Define output PNG file name (WebP not reliably supported in this environment).
                    string pngPath = Path.Combine(outputDir, $"image_{imageIndex}.png");

                    // Save as PNG.
                    gifImage.Save(pngPath, ImageFormat.Png);

                    // Validate that the PNG file was created.
                    if (!File.Exists(pngPath))
                        throw new InvalidOperationException($"Failed to create PNG file: {pngPath}");
                }
            }

            imageIndex++;
        }

        // Final validation: at least one PNG file should exist.
        if (imageIndex == 0)
            throw new InvalidOperationException("No GIF images were found in the document.");

        // All done.
    }
}
