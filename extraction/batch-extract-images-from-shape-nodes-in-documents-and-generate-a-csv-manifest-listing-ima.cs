using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a sample document that contains a few images.
        const string inputPath = "sample.docx";
        CreateSampleDocument(inputPath);

        // Load the document that we just created.
        Document doc = new Document(inputPath);

        // Prepare an output folder for the extracted images and the CSV manifest.
        const string outputDir = "output";
        Directory.CreateDirectory(outputDir);

        // Prepare CSV lines – header first.
        List<string> csvLines = new List<string>
        {
            "ImageFileName,SourceDocument"
        };

        // Enumerate all shape nodes that actually contain an image.
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        foreach (Shape shape in shapes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                // Determine a proper file extension for the image type.
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string imageFileName = $"image_{imageIndex}{extension}";
                string imagePath = Path.Combine(outputDir, imageFileName);

                // Save the image to the file system.
                shape.ImageData.Save(imagePath);

                // Add an entry to the CSV manifest.
                csvLines.Add($"{imageFileName},{inputPath}");
                imageIndex++;
            }
        }

        // Write the CSV manifest file.
        string csvPath = Path.Combine(outputDir, "manifest.csv");
        File.WriteAllLines(csvPath, csvLines);

        // Validate that the manifest was created.
        if (!File.Exists(csvPath))
            throw new InvalidOperationException("CSV manifest was not created.");
    }

    // Helper method that creates a simple DOCX file with two identical PNG images.
    private static void CreateSampleDocument(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // A 1x1 pixel PNG image encoded as Base64.
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9yhl4AAAAASUVORK5CYII=";
        byte[] pngBytes = Convert.FromBase64String(base64Png);

        // Insert the first image.
        using (MemoryStream ms = new MemoryStream(pngBytes))
        {
            builder.InsertImage(ms);
        }

        // Insert a second image.
        using (MemoryStream ms = new MemoryStream(pngBytes))
        {
            builder.InsertImage(ms);
        }

        // Save the document to the specified path.
        doc.Save(filePath);
    }
}
