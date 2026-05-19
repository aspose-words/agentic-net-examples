using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Base directory for the demo
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "ExtractionDemo");
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string imageDir = Path.Combine(baseDir, "ExtractedImages");
        string manifestPath = Path.Combine(baseDir, "manifest.csv");

        // Ensure directories exist
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(imageDir);

        // Create sample documents containing images
        CreateSampleDocument(Path.Combine(inputDir, "Doc1.docx"));
        CreateSampleDocument(Path.Combine(inputDir, "Doc2.docx"));

        // Prepare CSV manifest header
        List<string> csvLines = new List<string>();
        csvLines.Add("DocumentPath,ImageFileName");

        // Process each document in the input folder
        foreach (string docPath in Directory.GetFiles(inputDir, "*.docx"))
        {
            Document doc = new Document(docPath);
            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;

            foreach (Shape shape in shapes.OfType<Shape>())
            {
                if (shape.HasImage)
                {
                    // Determine file extension based on image type
                    string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                    // Create deterministic image file name
                    string imageFileName = $"{Path.GetFileNameWithoutExtension(docPath)}_img{imageIndex}{extension}";
                    string imageFullPath = Path.Combine(imageDir, imageFileName);
                    // Save the image to the file system
                    shape.ImageData.Save(imageFullPath);
                    // Add entry to CSV manifest
                    csvLines.Add($"\"{docPath}\",\"{imageFileName}\"");
                    imageIndex++;
                }
            }
        }

        // Write the CSV manifest
        File.WriteAllLines(manifestPath, csvLines);

        // Validate that the manifest was created
        if (!File.Exists(manifestPath))
            throw new InvalidOperationException("CSV manifest was not created.");

        // Indicate completion (no interactive prompts)
        Console.WriteLine("Image extraction and manifest generation completed.");
    }

    // Creates a simple document with two identical PNG images inserted inline
    private static void CreateSampleDocument(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Small 1x1 pixel PNG image encoded in base64
        string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9yhl4AAAAASUVORK5CYII=";
        byte[] pngBytes = Convert.FromBase64String(base64Png);

        using (MemoryStream ms = new MemoryStream(pngBytes))
        {
            builder.InsertImage(ms);
        }

        builder.Writeln(); // Add a line break between images

        using (MemoryStream ms = new MemoryStream(pngBytes))
        {
            builder.InsertImage(ms);
        }

        // Save the document to the specified path
        doc.Save(filePath);
    }
}
