using System;
using System.IO;
using System.IO.Compression;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    // Entry point
    public static void Main()
    {
        // Define deterministic paths
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "BatchImageProcessing");
        string inputDocsDir = Path.Combine(baseDir, "InputDocs");
        string extractedImagesDir = Path.Combine(baseDir, "ExtractedImages");
        string zipPath = Path.Combine(baseDir, "ImagesArchive.zip");
        string sampleImagePath = Path.Combine(baseDir, "sample.png");

        // Ensure clean environment
        if (Directory.Exists(baseDir))
            Directory.Delete(baseDir, true);
        Directory.CreateDirectory(inputDocsDir);
        Directory.CreateDirectory(extractedImagesDir);

        // -------------------------------------------------
        // 1. Create a deterministic sample image (sample.png)
        // -------------------------------------------------
        CreateSampleImage(sampleImagePath);

        // -------------------------------------------------
        // 2. Create sample DOCX files that contain the image
        // -------------------------------------------------
        const int docCount = 3;
        for (int i = 1; i <= docCount; i++)
        {
            string docPath = Path.Combine(inputDocsDir, $"Document{i}.docx");
            CreateDocumentWithImage(docPath, sampleImagePath);
        }

        // -------------------------------------------------
        // 3. Batch process DOCX files: extract images
        // -------------------------------------------------
        int totalExtracted = 0;
        foreach (string docFile in Directory.GetFiles(inputDocsDir, "*.docx"))
        {
            totalExtracted += ExtractImagesFromDocument(docFile, extractedImagesDir);
        }

        if (totalExtracted == 0)
            throw new InvalidOperationException("No images were extracted from the documents.");

        // -------------------------------------------------
        // 4. Create a ZIP archive of extracted images
        // -------------------------------------------------
        CreateZipArchive(extractedImagesDir, zipPath);

        // Validate ZIP creation
        if (!File.Exists(zipPath))
            throw new InvalidOperationException("Failed to create the ZIP archive.");

        // Example completed successfully
        Console.WriteLine($"Processed {docCount} documents, extracted {totalExtracted} images, and created ZIP at: {zipPath}");
    }

    // Creates a simple white PNG image with a black rectangle for deterministic content
    private static void CreateSampleImage(string filePath)
    {
        const int width = 200;
        const int height = 100;

        // Use Aspose.Drawing types explicitly
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height);
        Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap);
        graphics.Clear(Aspose.Drawing.Color.White);

        using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Black, 3))
        {
            graphics.DrawRectangle(pen, 10, 10, width - 20, height - 20);
        }

        // Save the image
        bitmap.Save(filePath);

        // Clean up
        graphics.Dispose();
        bitmap.Dispose();
    }

    // Creates a DOCX file and inserts the provided image
    private static void CreateDocumentWithImage(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln($"Document generated on {DateTime.Now}");
        builder.InsertImage(imagePath);
        doc.Save(docPath);
    }

    // Extracts all images from a DOCX file and saves them to the target folder
    private static int ExtractImagesFromDocument(string docPath, string targetFolder)
    {
        Document doc = new Document(docPath);
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;

        foreach (Shape shape in shapes)
        {
            if (shape.HasImage)
            {
                string imageExt = shape.ImageData.ImageType.ToString().ToLower(); // e.g., png, jpeg
                string imageFileName = $"{Path.GetFileNameWithoutExtension(docPath)}_image{++imageIndex}.{imageExt}";
                string imageFullPath = Path.Combine(targetFolder, imageFileName);
                shape.ImageData.Save(imageFullPath);
            }
        }

        return imageIndex; // number of images extracted from this document
    }

    // Creates a ZIP archive containing all files from sourceFolder
    private static void CreateZipArchive(string sourceFolder, string zipFilePath)
    {
        // Ensure the destination zip does not already exist
        if (File.Exists(zipFilePath))
            File.Delete(zipFilePath);

        ZipFile.CreateFromDirectory(sourceFolder, zipFilePath, CompressionLevel.Optimal, false);
    }
}
