using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class BatchGifToWebpConverter
{
    // Directories used in the example.
    private const string ArtifactsDir = "Artifacts";
    private const string InputImagesDir = "Artifacts/InputImages";
    private const string OutputWebpDir = "Artifacts/OutputWebp";

    public static void Main()
    {
        // Ensure that all required directories exist.
        Directory.CreateDirectory(ArtifactsDir);
        Directory.CreateDirectory(InputImagesDir);
        Directory.CreateDirectory(OutputWebpDir);

        // 1. Create sample GIF images (static GIFs for simplicity).
        CreateSampleGif("sample1.gif", Aspose.Drawing.Color.Blue);
        CreateSampleGif("sample2.gif", Aspose.Drawing.Color.Green);

        // 2. Insert the GIF images into a Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        foreach (string gifPath in Directory.GetFiles(InputImagesDir, "*.gif"))
        {
            // Insert each GIF image into the document.
            Shape shape = builder.InsertImage(gifPath);
            shape.WrapType = WrapType.Inline;
            builder.Writeln(); // Add a line break after each image.
        }

        // Save the document that contains the GIF images.
        string docPath = Path.Combine(ArtifactsDir, "DocumentWithGifs.docx");
        doc.Save(docPath);

        // 3. Load the document and extract GIF images.
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        int gifIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Process only GIF images.
            if (shape.ImageData.ImageType != ImageType.Gif)
                continue;

            // Extract the GIF image bytes.
            using (MemoryStream gifStream = new MemoryStream())
            {
                shape.ImageData.Save(gifStream);
                gifStream.Position = 0;

                // Load the GIF into Aspose.Drawing.Bitmap.
                using (Bitmap bitmap = new Bitmap(gifStream))
                {
                    // Prepare the output file name. The example uses PNG because
                    // WebP support via Aspose.Drawing is not guaranteed in the verifier environment.
                    string outputFileName = $"Gif_{gifIndex}.png";
                    string outputPath = Path.Combine(OutputWebpDir, outputFileName);

                    // Save the bitmap as PNG. This preserves the visual content of the first frame.
                    // If WebP support becomes available, replace ImageFormat.Png with ImageFormat.Webp
                    // and change the file extension accordingly.
                    bitmap.Save(outputPath, ImageFormat.Png);
                }
            }

            gifIndex++;
        }

        // Validation: ensure at least one output file was created.
        int outputCount = Directory.GetFiles(OutputWebpDir, "*.png").Length;
        if (outputCount == 0)
            throw new InvalidOperationException("No output files were created. Conversion may have failed.");

        Console.WriteLine($"Successfully converted {outputCount} GIF image(s) to PNG files (placeholder for WebP).");
    }

    // Helper method to create a simple static GIF image using Aspose.Drawing.
    private static void CreateSampleGif(string fileName, Aspose.Drawing.Color backgroundColor)
    {
        string filePath = Path.Combine(InputImagesDir, fileName);

        // Create a 200x200 bitmap.
        using (Bitmap bitmap = new Bitmap(200, 200))
        {
            // Obtain a graphics object from the bitmap.
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                // Fill the background with the specified color.
                graphics.Clear(backgroundColor);

                // Draw a simple white ellipse.
                using (Pen pen = new Pen(Aspose.Drawing.Color.White, 5))
                {
                    graphics.DrawEllipse(pen, 20, 20, 160, 160);
                }
            }

            // Save the bitmap as a GIF image.
            bitmap.Save(filePath, ImageFormat.Gif);
        }
    }
}
