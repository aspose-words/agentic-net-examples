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
        // Base folder for the example.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "ImageBatchExample");
        string inputDocsDir = Path.Combine(baseDir, "InputDocs");
        string outputImagesDir = Path.Combine(baseDir, "ExtractedImages");

        // Ensure clean environment.
        if (Directory.Exists(baseDir))
            Directory.Delete(baseDir, true);
        Directory.CreateDirectory(inputDocsDir);
        Directory.CreateDirectory(outputImagesDir);

        // Create sample images of different formats.
        string[] imageFiles = CreateSampleImages(baseDir);

        // Create one or more sample DOCX files that contain the images.
        CreateSampleDocument(Path.Combine(inputDocsDir, "Sample1.docx"), imageFiles);
        CreateSampleDocument(Path.Combine(inputDocsDir, "Sample2.docx"), imageFiles);

        // Batch extract images from all DOC/DOCX files in the input folder.
        int totalExtracted = 0;
        foreach (string docPath in Directory.GetFiles(inputDocsDir, "*.*", SearchOption.TopDirectoryOnly)
                                            .Where(f => f.EndsWith(".doc", StringComparison.OrdinalIgnoreCase) ||
                                                        f.EndsWith(".docx", StringComparison.OrdinalIgnoreCase)))
        {
            Document doc = new Document(docPath);
            NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;

            foreach (Shape shape in shapeNodes.OfType<Shape>())
            {
                if (!shape.HasImage)
                    continue;

                // Determine file extension based on the image type stored in the shape.
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string formatFolder = Path.Combine(outputImagesDir, extension.TrimStart('.'));
                Directory.CreateDirectory(formatFolder);

                string outputFileName = $"{Path.GetFileNameWithoutExtension(docPath)}_{imageIndex}{extension}";
                string outputPath = Path.Combine(formatFolder, outputFileName);

                shape.ImageData.Save(outputPath);
                imageIndex++;
                totalExtracted++;
            }

            // Validation: ensure at least one image was extracted from this document.
            if (imageIndex == 0)
                throw new InvalidOperationException($"No images were extracted from document '{docPath}'.");
        }

        // Final validation: ensure the batch produced at least one image.
        if (totalExtracted == 0)
            throw new InvalidOperationException("Batch extraction completed but no images were found.");

        // Example completed successfully.
        Console.WriteLine($"Extraction finished. Total images extracted: {totalExtracted}");
    }

    // Creates deterministic sample images (PNG, JPEG, BMP, GIF) and returns their file paths.
    private static string[] CreateSampleImages(string baseDir)
    {
        string[] paths = new string[4];
        // PNG
        string pngPath = Path.Combine(baseDir, "sample.png");
        using (Bitmap bmp = new Bitmap(100, 100))
        using (Graphics g = Graphics.FromImage(bmp))
        {
            g.Clear(Aspose.Drawing.Color.LightBlue);
            bmp.Save(pngPath);
        }
        paths[0] = pngPath;

        // JPEG
        string jpgPath = Path.Combine(baseDir, "sample.jpg");
        using (Bitmap bmp = new Bitmap(120, 80))
        using (Graphics g = Graphics.FromImage(bmp))
        {
            g.Clear(Aspose.Drawing.Color.LightCoral);
            bmp.Save(jpgPath);
        }
        paths[1] = jpgPath;

        // BMP
        string bmpPath = Path.Combine(baseDir, "sample.bmp");
        using (Bitmap bmp = new Bitmap(80, 120))
        using (Graphics g = Graphics.FromImage(bmp))
        {
            g.Clear(Aspose.Drawing.Color.LightGreen);
            bmp.Save(bmpPath);
        }
        paths[2] = bmpPath;

        // GIF
        string gifPath = Path.Combine(baseDir, "sample.gif");
        using (Bitmap bmp = new Bitmap(90, 90))
        using (Graphics g = Graphics.FromImage(bmp))
        {
            g.Clear(Aspose.Drawing.Color.LightYellow);
            bmp.Save(gifPath);
        }
        paths[3] = gifPath;

        return paths;
    }

    // Creates a DOCX document and inserts each image from the provided list.
    private static void CreateSampleDocument(string docPath, string[] imageFiles)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        foreach (string imgPath in imageFiles)
        {
            // Insert image inline.
            builder.InsertImage(imgPath);
            builder.Writeln(); // Add a line break between images.
        }

        doc.Save(docPath);
    }
}
