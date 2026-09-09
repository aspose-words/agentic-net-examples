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

        // 1. Create a sample GIF image (single‑frame is sufficient for the demo)
        string gifPath = Path.Combine(artifactsDir, "sample.gif");
        CreateSampleGif(gifPath);

        // 2. Insert the GIF into a Word document
        string docPath = Path.Combine(artifactsDir, "DocWithGif.docx");
        InsertGifIntoDocument(gifPath, docPath);

        // 3. Extract the GIF from the document and split it into PNG frames
        ExtractGifFrames(docPath, artifactsDir);
    }

    private static void CreateSampleGif(string filePath)
    {
        // Create a 100x100 bitmap, fill it with a color and save as GIF
        using (Bitmap bitmap = new Bitmap(100, 100))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(Color.CornflowerBlue);
            bitmap.Save(filePath, ImageFormat.Gif);
        }
    }

    private static void InsertGifIntoDocument(string gifFile, string docFile)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        // Insert the GIF image
        builder.InsertImage(gifFile);
        // Save the document
        doc.Save(docFile);
    }

    private static void ExtractGifFrames(string docFile, string outputDir)
    {
        Document doc = new Document(docFile);
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

        int gifIndex = 0;
        foreach (Shape shape in shapes.OfType<Shape>())
        {
            if (!shape.HasImage) continue;

            // Process only GIF images
            if (shape.ImageData.ImageType != ImageType.Gif) continue;

            // Save the extracted GIF to a temporary file
            string extractedGifPath = Path.Combine(outputDir, $"extracted_{gifIndex}.gif");
            shape.ImageData.Save(extractedGifPath);

            // Load the GIF using Aspose.Drawing.Image
            using (Image gifImage = Image.FromFile(extractedGifPath))
            {
                // Determine the number of frames (time dimension)
                int frameCount = gifImage.GetFrameCount(FrameDimension.Time);
                if (frameCount == 0) frameCount = 1; // fallback for single‑frame GIFs

                for (int i = 0; i < frameCount; i++)
                {
                    // Select the current frame
                    gifImage.SelectActiveFrame(FrameDimension.Time, i);

                    // Create a bitmap from the current frame and save as PNG
                    using (Bitmap frameBitmap = new Bitmap(gifImage))
                    {
                        string pngPath = Path.Combine(outputDir, $"gif_{gifIndex}_frame_{i}.png");
                        frameBitmap.Save(pngPath, ImageFormat.Png);
                    }
                }
            }

            gifIndex++;
        }

        // Validation: ensure at least one PNG was created
        int pngCount = Directory.GetFiles(outputDir, "*.png").Length;
        if (pngCount == 0)
            throw new InvalidOperationException("No PNG frames were generated from the GIF image.");
    }
}
