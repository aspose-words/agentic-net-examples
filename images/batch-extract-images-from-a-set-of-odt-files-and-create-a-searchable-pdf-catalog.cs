using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare folders
        string baseDir = Directory.GetCurrentDirectory();
        string inputFolder = Path.Combine(baseDir, "InputDocs");
        string imageFolder = Path.Combine(baseDir, "ExtractedImages");
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(imageFolder);

        // Create sample images that will be inserted into ODT files
        string sampleImage1 = Path.Combine(baseDir, "sample1.png");
        string sampleImage2 = Path.Combine(baseDir, "sample2.png");
        CreateSampleImage(sampleImage1, Aspose.Drawing.Color.LightBlue);
        CreateSampleImage(sampleImage2, Aspose.Drawing.Color.LightGreen);

        // Create a few ODT documents containing the sample images
        CreateSampleOdtDocument(Path.Combine(inputFolder, "Doc1.odt"), sampleImage1);
        CreateSampleOdtDocument(Path.Combine(inputFolder, "Doc2.odt"), sampleImage2);

        // Extract images from all ODT files in the input folder
        List<(string SourceDoc, string ImagePath)> extractedImages = new List<(string, string)>();
        foreach (string odtPath in Directory.GetFiles(inputFolder, "*.odt"))
        {
            Document doc = new Document(odtPath);
            NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;
            foreach (Shape shape in shapeNodes)
            {
                if (shape.HasImage)
                {
                    string imageFileName = $"{Path.GetFileNameWithoutExtension(odtPath)}-image{imageIndex}.png";
                    string imagePath = Path.Combine(imageFolder, imageFileName);
                    shape.ImageData.Save(imagePath);
                    extractedImages.Add((Path.GetFileName(odtPath), imagePath));
                    imageIndex++;
                }
            }
        }

        // Validate that at least one image was extracted
        if (extractedImages.Count == 0)
            throw new InvalidOperationException("No images were extracted from the ODT files.");

        // Create a searchable PDF catalog containing the extracted images
        Document catalog = new Document();
        DocumentBuilder builder = new DocumentBuilder(catalog);
        foreach (var (sourceDoc, imagePath) in extractedImages)
        {
            builder.Writeln($"Image extracted from {sourceDoc}:");
            builder.InsertImage(imagePath);
            builder.InsertBreak(BreakType.PageBreak);
        }

        string catalogPath = Path.Combine(baseDir, "Catalog.pdf");
        catalog.Save(catalogPath, SaveFormat.Pdf);

        // Validate that the PDF catalog was created
        if (!File.Exists(catalogPath))
            throw new InvalidOperationException("Failed to create the PDF catalog.");

        // Cleanup sample images (optional)
        // File.Delete(sampleImage1);
        // File.Delete(sampleImage2);
    }

    // Creates a simple PNG image with a solid background color
    private static void CreateSampleImage(string filePath, Aspose.Drawing.Color backgroundColor)
    {
        int width = 200;
        int height = 100;
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(backgroundColor);
            }
            bitmap.Save(filePath);
        }
    }

    // Creates an ODT document that contains a single image
    private static void CreateSampleOdtDocument(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln($"This document contains an image:");
        builder.InsertImage(imagePath);
        doc.Save(docPath, SaveFormat.Odt);
    }
}
