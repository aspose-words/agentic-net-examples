using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Newtonsoft.Json;

public class BatchImageExtraction
{
    public static void Main()
    {
        // Set up directories
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "BatchExtractionDemo");
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string outputDir = Path.Combine(baseDir, "ExtractedImages");
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create sample documents with images
        CreateSampleDocument(Path.Combine(inputDir, "Doc1.docx"), 2);
        CreateSampleDocument(Path.Combine(inputDir, "Doc2.docx"), 1);

        // Prepare manifest data
        List<string[]> manifestRows = new List<string[]>();
        manifestRows.Add(new[] { "DocumentName", "ImageFileName", "ImageSource" });

        // Process each document
        string[] docFiles = Directory.GetFiles(inputDir, "*.docx");
        foreach (string docPath in docFiles)
        {
            Document doc = new Document(docPath);
            NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
            var imageShapes = shapeNodes.OfType<Shape>().Where(s => s.HasImage).ToList();

            if (imageShapes.Count == 0)
            {
                continue; // No images in this document
            }

            int imageIndex = 0;
            foreach (Shape shape in imageShapes)
            {
                string docName = Path.GetFileName(docPath);
                string imageFileName = $"{Path.GetFileNameWithoutExtension(docName)}_Image{imageIndex}.png";
                string imagePath = Path.Combine(outputDir, imageFileName);

                // Save the image
                shape.ImageData.Save(imagePath);

                // Record manifest entry
                string imageSource = string.IsNullOrEmpty(shape.Name) ? "UnnamedShape" : shape.Name;
                manifestRows.Add(new[] { docName, imageFileName, imageSource });

                imageIndex++;
            }
        }

        // Write CSV manifest
        string manifestPath = Path.Combine(outputDir, "manifest.csv");
        File.WriteAllLines(manifestPath, manifestRows.Select(row => string.Join(",", row)));

        // Validation
        if (!File.Exists(manifestPath))
        {
            throw new InvalidOperationException("CSV manifest was not created.");
        }

        // Ensure at least one image was extracted
        string[] extractedImages = Directory.GetFiles(outputDir, "*.png");
        if (extractedImages.Length == 0)
        {
            throw new InvalidOperationException("No images were extracted.");
        }

        // Optional: Output a simple JSON summary (not required but demonstrates Newtonsoft.Json usage)
        var summary = new
        {
            DocumentsProcessed = docFiles.Length,
            ImagesExtracted = extractedImages.Length,
            ManifestFile = manifestPath
        };
        string jsonPath = Path.Combine(outputDir, "summary.json");
        File.WriteAllText(jsonPath, JsonConvert.SerializeObject(summary, Formatting.Indented));
    }

    private static void CreateSampleDocument(string filePath, int imageCount)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Base64-encoded 1x1 pixel PNG (transparent)
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9yhl4AAAAASUVORK5CYII=";
        byte[] pngBytes = Convert.FromBase64String(base64Png);

        for (int i = 0; i < imageCount; i++)
        {
            builder.Writeln($"Image {i + 1} in document.");
            using (MemoryStream ms = new MemoryStream(pngBytes))
            {
                builder.InsertImage(ms);
            }
        }

        doc.Save(filePath);
    }
}
