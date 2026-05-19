using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;

public class SepiaImageExtractor
{
    public static void Main()
    {
        // Prepare folders.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        string inputDocsDir = Path.Combine(artifactsDir, "InputDocs");
        string outputImagesDir = Path.Combine(artifactsDir, "SepiaImages");
        Directory.CreateDirectory(artifactsDir);
        Directory.CreateDirectory(inputDocsDir);
        Directory.CreateDirectory(outputImagesDir);

        // Create a deterministic sample PNG image.
        string samplePngPath = Path.Combine(artifactsDir, "sample.png");
        CreateSamplePng(samplePngPath);

        // Create a few sample Word documents that contain the PNG image.
        CreateSampleDocument(Path.Combine(inputDocsDir, "Doc1.docx"), samplePngPath);
        CreateSampleDocument(Path.Combine(inputDocsDir, "Doc2.docx"), samplePngPath);

        // Process each document: extract PNG images, apply sepia, and save.
        int totalExtracted = 0;
        foreach (string docPath in Directory.GetFiles(inputDocsDir, "*.docx"))
        {
            Document doc = new Document(docPath);
            NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;

            foreach (Shape shape in shapeNodes.OfType<Shape>())
            {
                if (!shape.HasImage)
                    continue;

                if (shape.ImageData.ImageType != ImageType.Png)
                    continue;

                // Save the original image to a memory stream.
                using (MemoryStream originalStream = new MemoryStream())
                {
                    shape.ImageData.Save(originalStream);
                    originalStream.Position = 0;

                    // Load the image with Aspose.Drawing.
                    using (Bitmap originalBitmap = new Bitmap(originalStream))
                    {
                        // Apply sepia tone.
                        using (Bitmap sepiaBitmap = ApplySepia(originalBitmap))
                        {
                            string outFileName = $"sepia_{Path.GetFileNameWithoutExtension(docPath)}_{imageIndex}.png";
                            string outPath = Path.Combine(outputImagesDir, outFileName);
                            sepiaBitmap.Save(outPath);
                            totalExtracted++;
                        }
                    }
                }

                imageIndex++;
            }
        }

        // Validate that at least one image was produced.
        if (totalExtracted == 0)
            throw new InvalidOperationException("No PNG images were extracted and processed.");

        // Example completed successfully.
        Console.WriteLine($"Sepia processing completed. {totalExtracted} image(s) saved to '{outputImagesDir}'.");
    }

    // Creates a simple 200x200 PNG with a solid background.
    private static void CreateSamplePng(string filePath)
    {
        using (Bitmap bitmap = new Bitmap(200, 200))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(Color.LightBlue);
            bitmap.Save(filePath);
        }
    }

    // Creates a Word document that inserts the specified image.
    private static void CreateSampleDocument(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath);
    }

    // Returns a new bitmap with a sepia tone applied.
    private static Bitmap ApplySepia(Bitmap source)
    {
        Bitmap result = new Bitmap(source.Width, source.Height);
        for (int y = 0; y < source.Height; y++)
        {
            for (int x = 0; x < source.Width; x++)
            {
                Color pixel = source.GetPixel(x, y);
                double r = pixel.R;
                double g = pixel.G;
                double b = pixel.B;

                int tr = (int)(0.393 * r + 0.769 * g + 0.189 * b);
                int tg = (int)(0.349 * r + 0.686 * g + 0.168 * b);
                int tb = (int)(0.272 * r + 0.534 * g + 0.131 * b);

                tr = Math.Min(255, tr);
                tg = Math.Min(255, tg);
                tb = Math.Min(255, tb);

                result.SetPixel(x, y, Color.FromArgb(tr, tg, tb));
            }
        }
        return result;
    }
}
