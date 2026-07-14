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
        // Prepare folders
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string inputImagePath = Path.Combine(artifactsDir, "sample.gif");
        string docPath = Path.Combine(artifactsDir, "DocumentWithGif.docx");

        // -----------------------------------------------------------------
        // 1. Create a simple GIF image (single‑frame is sufficient for demo)
        // -----------------------------------------------------------------
        int width = 200, height = 200;
        using (Bitmap bmp = new Bitmap(width, height))
        {
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.Clear(Color.White);
                g.DrawEllipse(new Pen(Color.Blue, 5), 10, 10, width - 20, height - 20);
            }
            // Save as GIF
            bmp.Save(inputImagePath, ImageFormat.Gif);
        }

        // --------------------------------------------------------------
        // 2. Insert the GIF into a Word document and save the document
        // --------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(inputImagePath);
        doc.Save(docPath);

        // --------------------------------------------------------------
        // 3. Reload the document and locate the GIF shape
        // --------------------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        Shape gifShape = null;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage && shape.ImageData.ImageType == ImageType.Gif)
            {
                gifShape = shape;
                break;
            }
        }

        if (gifShape == null)
            throw new InvalidOperationException("No GIF image found in the document.");

        // --------------------------------------------------------------
        // 4. Extract the GIF bytes and load it with Aspose.Drawing.Image
        // --------------------------------------------------------------
        using (MemoryStream gifStream = new MemoryStream())
        {
            gifShape.ImageData.Save(gifStream);
            gifStream.Position = 0;

            using (Image gifImage = Image.FromStream(gifStream))
            {
                // Determine the number of frames in the GIF
                int frameCount = gifImage.GetFrameCount(FrameDimension.Time);
                if (frameCount == 0) frameCount = 1; // fallback for non‑animated GIFs

                // --------------------------------------------------------------
                // 5. Save each frame as a separate PNG file
                // --------------------------------------------------------------
                for (int i = 0; i < frameCount; i++)
                {
                    gifImage.SelectActiveFrame(FrameDimension.Time, i);
                    string pngPath = Path.Combine(artifactsDir, $"frame_{i + 1}.png");
                    gifImage.Save(pngPath, ImageFormat.Png);
                }

                // Validation: ensure at least one PNG was created
                string[] pngFiles = Directory.GetFiles(artifactsDir, "frame_*.png");
                if (pngFiles.Length == 0)
                    throw new InvalidOperationException("No PNG frames were generated.");
            }
        }
    }
}
