using System;
using System.IO;
using System.IO.Compression;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Base directory for all example files
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "ExampleData");
        string inputDocsDir = Path.Combine(baseDir, "InputDocs");
        string extractedImagesDir = Path.Combine(baseDir, "ExtractedImages");
        string zipPath = Path.Combine(baseDir, "ImagesArchive.zip");

        // Ensure directories exist
        Directory.CreateDirectory(inputDocsDir);
        Directory.CreateDirectory(extractedImagesDir);

        // Create a deterministic sample image (input.png)
        string sampleImagePath = Path.Combine(baseDir, "input.png");
        CreateSampleImage(sampleImagePath, 200, 200);

        // Create sample DOCX files containing the sample image
        CreateSampleDocuments(inputDocsDir, sampleImagePath, 2);

        // Batch process DOCX files: extract images
        int extractedCount = ExtractImagesFromDocuments(inputDocsDir, extractedImagesDir);

        // Validate that at least one image was extracted
        if (extractedCount == 0)
            throw new InvalidOperationException("No images were extracted from the documents.");

        // Create a zip archive containing all extracted images
        if (File.Exists(zipPath))
            File.Delete(zipPath);
        ZipFile.CreateFromDirectory(extractedImagesDir, zipPath);

        // Validate zip creation
        if (!File.Exists(zipPath))
            throw new InvalidOperationException("Failed to create the zip archive.");

        // Example completed successfully
        Console.WriteLine($"Extracted {extractedCount} images and created archive at: {zipPath}");
    }

    private static void CreateSampleImage(string filePath, int width, int height)
    {
        // Create a white bitmap and save it as PNG
        Bitmap bitmap = new Bitmap(width, height);
        Graphics graphics = Graphics.FromImage(bitmap);
        graphics.Clear(Color.White);
        // Optionally draw something deterministic
        // (e.g., a simple rectangle)
        graphics.Dispose();
        bitmap.Save(filePath, ImageFormat.Png);
        bitmap.Dispose();
    }

    private static void CreateSampleDocuments(string outputDir, string imagePath, int documentCount)
    {
        for (int i = 1; i <= documentCount; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"Document {i} with sample image:");
            // Insert the sample image twice to have multiple images per document
            builder.InsertImage(imagePath);
            builder.InsertParagraph();
            builder.InsertImage(imagePath);
            string docPath = Path.Combine(outputDir, $"Sample{i}.docx");
            doc.Save(docPath);
        }
    }

    private static int ExtractImagesFromDocuments(string docsDir, string outputImagesDir)
    {
        int totalExtracted = 0;
        string[] docFiles = Directory.GetFiles(docsDir, "*.docx");
        foreach (string docFile in docFiles)
        {
            Document doc = new Document(docFile);
            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;
            foreach (Shape shape in shapes)
            {
                if (shape.HasImage)
                {
                    string imageFileName = $"{Path.GetFileNameWithoutExtension(docFile)}_image{imageIndex}.png";
                    string imagePath = Path.Combine(outputImagesDir, imageFileName);
                    shape.ImageData.Save(imagePath);
                    imageIndex++;
                    totalExtracted++;
                }
            }
        }
        return totalExtracted;
    }
}
