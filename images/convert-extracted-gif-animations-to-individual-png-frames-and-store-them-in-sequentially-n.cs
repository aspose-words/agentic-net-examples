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
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a sample GIF image (single‑frame for simplicity).
        // -----------------------------------------------------------------
        string gifPath = Path.Combine(outputDir, "sample.gif");
        using (Bitmap bmp = new Bitmap(200, 200))
        {
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.Clear(Color.Blue);
            }
            bmp.Save(gifPath, ImageFormat.Gif);
        }

        // -----------------------------------------------------------------
        // 2. Insert the GIF into a Word document.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(gifPath);
        string docPath = Path.Combine(outputDir, "DocumentWithGif.docx");
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 3. Extract the GIF image from the document.
        // -----------------------------------------------------------------
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        string extractedGifPath = Path.Combine(outputDir, "extracted.gif");
        bool gifFound = false;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage && shape.ImageData.ImageType == ImageType.Gif)
            {
                shape.ImageData.Save(extractedGifPath);
                gifFound = true;
                break; // Assume only one GIF for this example.
            }
        }

        if (!gifFound)
            throw new InvalidOperationException("No GIF image was found in the document.");

        // -----------------------------------------------------------------
        // 4. Split the GIF into individual PNG frames.
        // -----------------------------------------------------------------
        using (Image gifImage = Image.FromFile(extractedGifPath))
        {
            // FrameDimension.Time represents the animation frames of a GIF.
            FrameDimension dimension = new FrameDimension(gifImage.FrameDimensionsList[0]);
            int frameCount = gifImage.GetFrameCount(dimension);

            if (frameCount == 0)
                throw new InvalidOperationException("The extracted GIF contains no frames.");

            for (int i = 0; i < frameCount; i++)
            {
                gifImage.SelectActiveFrame(dimension, i);
                using (Bitmap frameBmp = new Bitmap(gifImage))
                {
                    string framePath = Path.Combine(outputDir, $"frame_{i + 1}.png");
                    frameBmp.Save(framePath, ImageFormat.Png);
                }
            }
        }

        // -----------------------------------------------------------------
        // 5. Validate that at least one PNG file was created.
        // -----------------------------------------------------------------
        string[] pngFiles = Directory.GetFiles(outputDir, "frame_*.png");
        if (pngFiles.Length == 0)
            throw new InvalidOperationException("No PNG frames were generated from the GIF.");

        // Example completed successfully.
        Console.WriteLine($"Generated {pngFiles.Length} PNG frame(s) in: {outputDir}");
    }
}
