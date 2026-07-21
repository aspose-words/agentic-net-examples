using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Ensure output directories exist
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // 1. Create a deterministic sample PNG image using Aspose.Drawing
        string sampleImagePath = Path.Combine(outputDir, "sample.png");
        CreateSamplePng(sampleImagePath, 200, 200);

        // 2. Insert the sample image into a new Word document
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        string docPath = Path.Combine(outputDir, "sample.docx");
        doc.Save(docPath);

        // 3. Extract all images from the document
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Determine file extension based on image type
            string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
            string extractedPath = Path.Combine(outputDir, $"extracted{imageIndex}{extension}");

            // Save the extracted image
            shape.ImageData.Save(extractedPath);

            // 4. If the image is a PNG, apply a color inversion filter
            if (shape.ImageData.ImageType == ImageType.Png)
            {
                string invertedPath = Path.Combine(outputDir, $"inverted{imageIndex}{extension}");
                InvertPngColors(extractedPath, invertedPath);
                if (!File.Exists(invertedPath))
                    throw new InvalidOperationException($"Inverted image was not created: {invertedPath}");
            }

            imageIndex++;
        }

        // Validation: at least one PNG image should have been processed
        if (imageIndex == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        // Example completed successfully
        Console.WriteLine("Image extraction and inversion completed.");
    }

    // Creates a simple PNG with a solid colored rectangle
    private static void CreateSamplePng(string filePath, int width, int height)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(Color.White);
            // Draw a red rectangle
            graphics.FillRectangle(new SolidBrush(Color.Red), 20, 20, width - 40, height - 40);
            bitmap.Save(filePath);
        }
    }

    // Inverts colors of a PNG image and saves the result
    private static void InvertPngColors(string sourcePath, string destinationPath)
    {
        using (Bitmap bitmap = new Bitmap(sourcePath))
        {
            for (int y = 0; y < bitmap.Height; y++)
            {
                for (int x = 0; x < bitmap.Width; x++)
                {
                    Color original = bitmap.GetPixel(x, y);
                    Color inverted = Color.FromArgb(255 - original.R, 255 - original.G, 255 - original.B);
                    bitmap.SetPixel(x, y, inverted);
                }
            }
            bitmap.Save(destinationPath);
        }
    }
}
