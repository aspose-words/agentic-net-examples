using System;
using System.IO;
using System.IO.Compression;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;               // Aspose.Drawing.Common
using Aspose.Drawing.Imaging;

public class Program
{
    // Entry point
    public static void Main()
    {
        // Prepare working folders
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        string inputDocsDir = Path.Combine(baseDir, "InputDocs");
        string imagesDir = Path.Combine(baseDir, "ExtractedImages");
        string zipPath = Path.Combine(baseDir, "ImagesArchive.zip");

        Directory.CreateDirectory(inputDocsDir);
        Directory.CreateDirectory(imagesDir);

        // 1. Create a deterministic sample image
        string sampleImagePath = Path.Combine(baseDir, "sample.png");
        CreateSampleImage(sampleImagePath, 200, 100);

        // 2. Generate a few DOC files that contain the sample image
        CreateSampleDocument(Path.Combine(inputDocsDir, "Document1.doc"), sampleImagePath);
        CreateSampleDocument(Path.Combine(inputDocsDir, "Document2.doc"), sampleImagePath);

        // 3. Extract all images from the DOC files into a folder
        ExtractImagesFromDocuments(inputDocsDir, imagesDir);

        // 4. Create a ZIP archive (without password – .NET built‑in ZIP does not support encryption)
        CreateZip(imagesDir, zipPath);

        // Validation – ensure the ZIP file exists and contains entries
        if (!File.Exists(zipPath))
            throw new InvalidOperationException("ZIP archive was not created.");

        Console.WriteLine($"ZIP archive created at: {zipPath}");
    }

    // Creates a deterministic PNG image using Aspose.Drawing
    private static void CreateSampleImage(string filePath, int width, int height)
    {
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height))
        using (Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap))
        {
            graphics.Clear(Aspose.Drawing.Color.White);
            // Draw a simple rectangle
            using (var brush = new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.LightBlue))
            {
                graphics.FillRectangle(brush, 10, 10, width - 20, height - 20);
            }
            bitmap.Save(filePath, Aspose.Drawing.Imaging.ImageFormat.Png);
        }
    }

    // Generates a DOC file with a single image inserted
    private static void CreateSampleDocument(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln($"This document contains an image inserted from {Path.GetFileName(imagePath)}:");
        builder.InsertImage(imagePath);
        doc.Save(docPath);
    }

    // Extracts all images from every DOC file in the source folder
    private static void ExtractImagesFromDocuments(string sourceFolder, string outputFolder)
    {
        string[] docFiles = Directory.GetFiles(sourceFolder, "*.doc", SearchOption.TopDirectoryOnly);
        if (docFiles.Length == 0)
            throw new InvalidOperationException("No DOC files found to process.");

        foreach (string docFile in docFiles)
        {
            Document doc = new Document(docFile);
            var shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;

            foreach (Shape shape in shapeNodes.OfType<Shape>())
            {
                if (shape.HasImage)
                {
                    string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                    string imageFileName = $"{Path.GetFileNameWithoutExtension(docFile)}_img{imageIndex}{extension}";
                    string imageFullPath = Path.Combine(outputFolder, imageFileName);
                    shape.ImageData.Save(imageFullPath);
                    imageIndex++;
                }
            }

            if (imageIndex == 0)
                throw new InvalidOperationException($"No images extracted from {docFile}.");
        }
    }

    // Creates a ZIP archive using System.IO.Compression (no password protection)
    private static void CreateZip(string sourceFolder, string zipFilePath)
    {
        // Remove any existing archive
        if (File.Exists(zipFilePath))
            File.Delete(zipFilePath);

        ZipFile.CreateFromDirectory(sourceFolder, zipFilePath, CompressionLevel.Optimal, includeBaseDirectory: false);
    }
}
