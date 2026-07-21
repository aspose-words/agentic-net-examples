using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Base directories for the example.
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string imageOutputDir = Path.Combine(baseDir, "ExtractedImages");
        string manifestPath = Path.Combine(baseDir, "manifest.csv");

        // Ensure clean folders.
        if (Directory.Exists(inputDir)) Directory.Delete(inputDir, true);
        if (Directory.Exists(imageOutputDir)) Directory.Delete(imageOutputDir, true);
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(imageOutputDir);

        // Create a few sample documents each containing an image.
        CreateSampleDocuments(inputDir);

        // Prepare CSV manifest lines.
        List<string> csvLines = new List<string>();
        csvLines.Add("ImageFileName,SourceDocument");

        // Process each DOCX file in the input folder.
        string[] docFiles = Directory.GetFiles(inputDir, "*.docx");
        foreach (string docPath in docFiles)
        {
            Document doc = new Document(docPath);
            NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;

            foreach (Shape shape in shapeNodes.OfType<Shape>())
            {
                if (!shape.HasImage) continue;

                // Determine deterministic image file name.
                string imageExtension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string imageFileName = $"{Path.GetFileNameWithoutExtension(docPath)}_Image{imageIndex}{imageExtension}";
                string imageFullPath = Path.Combine(imageOutputDir, imageFileName);

                // Save the image.
                shape.ImageData.Save(imageFullPath);
                if (!File.Exists(imageFullPath))
                    throw new InvalidOperationException($"Failed to save image '{imageFullPath}'.");

                // Add entry to manifest.
                csvLines.Add($"{imageFileName},{docPath}");
                imageIndex++;
            }

            // Ensure at least one image was extracted from this document.
            if (imageIndex == 0)
                throw new InvalidOperationException($"No images found in document '{docPath}'.");
        }

        // Write the CSV manifest.
        File.WriteAllLines(manifestPath, csvLines, Encoding.UTF8);
        if (!File.Exists(manifestPath))
            throw new InvalidOperationException("CSV manifest was not created.");

        // Example completed successfully.
        Console.WriteLine("Image extraction and manifest generation completed.");
    }

    // Creates three sample DOCX files each containing the same embedded PNG image.
    private static void CreateSampleDocuments(string folder)
    {
        // A 1x1 pixel transparent PNG (base64 encoded).
        const string pngBase64 =
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9yhl4AAAAASUVORK5CYII=";
        byte[] pngBytes = Convert.FromBase64String(pngBase64);

        for (int i = 0; i < 3; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            using (MemoryStream imageStream = new MemoryStream(pngBytes))
            {
                builder.InsertImage(imageStream);
            }
            string docPath = Path.Combine(folder, $"SampleDoc{i}.docx");
            doc.Save(docPath);
            if (!File.Exists(docPath))
                throw new InvalidOperationException($"Failed to create sample document '{docPath}'.");
        }
    }
}
