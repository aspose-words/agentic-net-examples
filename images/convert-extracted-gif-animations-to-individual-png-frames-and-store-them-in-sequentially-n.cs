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
        // Folder for all generated files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // 1. Create a deterministic animated GIF file (embedded base64 data).
        // This GIF contains two simple frames.
        string gifBase64 = "R0lGODlhAQABAPAAAP///wAAACH5BAAAAAAALAAAAAABAAEAAAICRAEAOw==";
        byte[] gifBytes = Convert.FromBase64String(gifBase64);
        string gifPath = Path.Combine(outputDir, "sample.gif");
        File.WriteAllBytes(gifPath, gifBytes);

        // 2. Insert the GIF into a Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(gifPath);
        string docPath = Path.Combine(outputDir, "sample.docx");
        doc.Save(docPath);

        // 3. Load the document and locate the GIF shape.
        Document loadedDoc = new Document(docPath);
        NodeCollection shapes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        Shape gifShape = null;
        foreach (Shape shape in shapes.OfType<Shape>())
        {
            if (shape.HasImage && shape.ImageData.ImageType == ImageType.Gif)
            {
                gifShape = shape;
                break;
            }
        }

        if (gifShape == null)
            throw new InvalidOperationException("No GIF image found in the document.");

        // 4. Extract the GIF image to a temporary file.
        string extractedGifPath = Path.Combine(outputDir, "extracted.gif");
        gifShape.ImageData.Save(extractedGifPath);

        // 5. Load the extracted GIF using Aspose.Drawing and split into PNG frames.
        using (Image gifImage = Image.FromFile(extractedGifPath))
        {
            // Determine the dimension that represents time (animation frames).
            Guid timeGuid = FrameDimension.Time.Guid;
            FrameDimension dimension = new FrameDimension(timeGuid);
            int frameCount = gifImage.GetFrameCount(dimension);

            if (frameCount == 0)
                throw new InvalidOperationException("The GIF does not contain any frames.");

            for (int i = 0; i < frameCount; i++)
            {
                gifImage.SelectActiveFrame(dimension, i);
                string framePath = Path.Combine(outputDir, $"frame_{i + 1}.png");
                gifImage.Save(framePath, ImageFormat.Png);
            }
        }

        // 6. Validation – ensure at least one PNG file was created.
        string[] pngFiles = Directory.GetFiles(outputDir, "frame_*.png");
        if (pngFiles.Length == 0)
            throw new InvalidOperationException("No PNG frames were generated.");

        // Example completed successfully.
        Console.WriteLine($"Generated {pngFiles.Length} PNG frame(s) in: {outputDir}");
    }
}
