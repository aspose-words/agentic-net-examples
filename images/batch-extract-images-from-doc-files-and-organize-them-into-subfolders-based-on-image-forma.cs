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
        // Base directories
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "BatchImageExtraction");
        string imagesInputDir = Path.Combine(baseDir, "InputImages");
        string docPath = Path.Combine(baseDir, "SampleDocument.docx");
        string outputDir = Path.Combine(baseDir, "ExtractedImages");

        // Ensure clean environment
        if (Directory.Exists(baseDir))
            Directory.Delete(baseDir, true);
        Directory.CreateDirectory(imagesInputDir);
        Directory.CreateDirectory(outputDir);

        // Create sample images of different formats
        CreateSampleImage(Path.Combine(imagesInputDir, "sample_png.png"), 200, 100, Aspose.Drawing.Color.LightBlue);
        CreateSampleImage(Path.Combine(imagesInputDir, "sample_jpeg.jpg"), 150, 150, Aspose.Drawing.Color.LightCoral);
        CreateSampleImage(Path.Combine(imagesInputDir, "sample_bmp.bmp"), 120, 180, Aspose.Drawing.Color.LightGreen);
        CreateSampleImage(Path.Combine(imagesInputDir, "sample_gif.gif"), 180, 120, Aspose.Drawing.Color.LightYellow);

        // Build a document and insert the images
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Document with multiple images:");
        builder.InsertImage(Path.Combine(imagesInputDir, "sample_png.png"));
        builder.InsertParagraph();
        builder.InsertImage(Path.Combine(imagesInputDir, "sample_jpeg.jpg"));
        builder.InsertParagraph();
        builder.InsertImage(Path.Combine(imagesInputDir, "sample_bmp.bmp"));
        builder.InsertParagraph();
        builder.InsertImage(Path.Combine(imagesInputDir, "sample_gif.gif"));
        doc.Save(docPath);

        // Load the document for extraction
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        int extractedCount = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Determine image type and corresponding extension
            ImageType imgType = shape.ImageData.ImageType;
            string extension = FileFormatUtil.ImageTypeToExtension(imgType); // includes leading dot
            string formatFolderName = extension.TrimStart('.').ToLowerInvariant();

            // Create subfolder for this image format
            string formatFolderPath = Path.Combine(outputDir, formatFolderName);
            Directory.CreateDirectory(formatFolderPath);

            // Build unique file name
            string imageFileName = $"Image_{extractedCount + 1}{extension}";
            string fullPath = Path.Combine(formatFolderPath, imageFileName);

            // Save the image
            shape.ImageData.Save(fullPath);
            extractedCount++;
        }

        // Validation: ensure at least one image was extracted
        if (extractedCount == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        // Optional: verify each format folder contains at least one file
        var formatFolders = Directory.GetDirectories(outputDir);
        foreach (var folder in formatFolders)
        {
            if (!Directory.EnumerateFiles(folder).Any())
                throw new InvalidOperationException($"Expected images in folder '{folder}' but none were found.");
        }

        // Execution finished
        Console.WriteLine($"Extraction complete. {extractedCount} images saved to '{outputDir}'.");
    }

    // Helper method to create a deterministic sample image
    private static void CreateSampleImage(string filePath, int width, int height, Aspose.Drawing.Color backgroundColor)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(backgroundColor);
            bitmap.Save(filePath);
        }
    }
}
