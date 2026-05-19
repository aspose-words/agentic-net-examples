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
        // Prepare folders
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        Directory.CreateDirectory(workDir);
        string outputDir = Path.Combine(workDir, "Frames");
        Directory.CreateDirectory(outputDir);

        // 1. Create a sample GIF image (single‑frame for demonstration)
        string gifPath = Path.Combine(workDir, "sample.gif");
        CreateSampleGif(gifPath);

        // 2. Insert the GIF into a Word document
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(gifPath);
        string docPath = Path.Combine(workDir, "DocumentWithGif.docx");
        doc.Save(docPath);

        // 3. Load the document and locate the shape that holds the GIF
        Document loadedDoc = new Document(docPath);
        Shape gifShape = null;
        foreach (Shape shape in loadedDoc.GetChildNodes(NodeType.Shape, true))
        {
            if (shape.HasImage && shape.ImageData.ImageType == ImageType.Gif)
            {
                gifShape = shape;
                break;
            }
        }

        if (gifShape == null)
            throw new InvalidOperationException("No GIF image found in the document.");

        // 4. Extract the GIF bytes to a memory stream
        using (MemoryStream gifStream = new MemoryStream())
        {
            gifShape.ImageData.Save(gifStream);
            gifStream.Position = 0; // Reset before reading

            // 5. Load the GIF with Aspose.Drawing and split into frames
            using (Image gifImage = Image.FromStream(gifStream))
            {
                // FrameDimension.Time is used for animated GIFs; fallback to 0 if unavailable
                FrameDimension dimension = new FrameDimension(gifImage.FrameDimensionsList[0]);
                int frameCount = gifImage.GetFrameCount(dimension);

                if (frameCount == 0)
                    throw new InvalidOperationException("The GIF contains no frames.");

                for (int i = 0; i < frameCount; i++)
                {
                    gifImage.SelectActiveFrame(dimension, i);
                    using (Bitmap frameBitmap = new Bitmap(gifImage))
                    {
                        string framePath = Path.Combine(outputDir, $"frame_{i + 1}.png");
                        frameBitmap.Save(framePath, ImageFormat.Png);
                    }
                }

                // Validation: ensure at least one PNG was written
                int pngCount = Directory.GetFiles(outputDir, "*.png").Length;
                if (pngCount == 0)
                    throw new InvalidOperationException("No PNG frames were extracted.");
            }
        }

        // Clean up (optional)
        // Console.WriteLine($"Extracted {Directory.GetFiles(outputDir, "*.png").Length} frame(s) to '{outputDir}'.");
    }

    private static void CreateSampleGif(string filePath)
    {
        // Create a simple bitmap and save it as a GIF file.
        using (Bitmap bitmap = new Bitmap(200, 200))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.LightCoral);
                graphics.DrawEllipse(new Pen(Color.White, 5), 20, 20, 160, 160);
            }

            bitmap.Save(filePath, ImageFormat.Gif);
        }
    }
}
