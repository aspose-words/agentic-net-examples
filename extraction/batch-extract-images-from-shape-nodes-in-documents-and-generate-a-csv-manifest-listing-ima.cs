using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Newtonsoft.Json;

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

        // Create sample documents with images
        CreateSampleDocument(Path.Combine(inputFolder, "Sample1.docx"));
        CreateSampleDocument(Path.Combine(inputFolder, "Sample2.docx"));

        // Prepare CSV manifest
        List<string> manifestLines = new List<string>();
        manifestLines.Add("Document,ImageFile"); // header

        // Process each document
        foreach (string docPath in Directory.GetFiles(inputFolder, "*.docx"))
        {
            Document doc = new Document(docPath);

            // Find all shape nodes that contain images
            var imageShapes = doc.GetChildNodes(NodeType.Shape, true)
                                 .OfType<Shape>()
                                 .Where(s => s.HasImage)
                                 .ToList();

            if (imageShapes.Count == 0)
                continue; // No images in this document

            for (int i = 0; i < imageShapes.Count; i++)
            {
                Shape shape = imageShapes[i];
                string imageFileName = $"{Path.GetFileNameWithoutExtension(docPath)}_Image{i}.png";
                string imagePath = Path.Combine(imageFolder, imageFileName);

                // Save the image data to a file
                shape.ImageData.Save(imagePath);

                // Add entry to manifest
                manifestLines.Add($"{Path.GetFileName(docPath)},{imageFileName}");
            }
        }

        // Write manifest CSV
        string manifestPath = Path.Combine(baseDir, "manifest.csv");
        File.WriteAllLines(manifestPath, manifestLines);

        // Validation
        if (!File.Exists(manifestPath))
            throw new InvalidOperationException("CSV manifest was not created.");

        // Ensure at least one image was extracted
        if (manifestLines.Count <= 1) // only header present
            throw new InvalidOperationException("No images were extracted from the documents.");

        // Example completed successfully
        Console.WriteLine("Image extraction and manifest generation completed.");
    }

    private static void CreateSampleDocument(string filePath)
    {
        // Simple 1x1 PNG image (transparent pixel) encoded in base64
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9yhl4AAAAASUVORK5CYII=";
        byte[] pngBytes = Convert.FromBase64String(base64Png);

        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert two images into the document
        using (MemoryStream ms = new MemoryStream(pngBytes))
        {
            builder.InsertImage(ms);
        }

        using (MemoryStream ms = new MemoryStream(pngBytes))
        {
            builder.InsertImage(ms);
        }

        doc.Save(filePath);
    }
}
