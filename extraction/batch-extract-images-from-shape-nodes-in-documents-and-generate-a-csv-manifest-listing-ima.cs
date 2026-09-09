using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare folders.
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string outputDir = Path.Combine(baseDir, "ExtractedImages");
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create a sample document with two embedded images.
        string sampleDocPath = Path.Combine(inputDir, "SampleDocument.docx");
        CreateSampleDocumentWithImages(sampleDocPath);

        // Prepare CSV manifest.
        List<string> csvLines = new List<string>();
        csvLines.Add("ImageFile,SourceDocument");

        // Process each DOCX file in the input folder.
        foreach (string docPath in Directory.GetFiles(inputDir, "*.docx"))
        {
            Document doc = new Document(docPath);
            var shapes = doc.GetChildNodes(NodeType.Shape, true).OfType<Shape>();
            int imageIndex = 0;

            foreach (Shape shape in shapes)
            {
                if (!shape.HasImage)
                    continue;

                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string imageFileName = $"{Path.GetFileNameWithoutExtension(docPath)}_Image{imageIndex}{extension}";
                string imageFullPath = Path.Combine(outputDir, imageFileName);

                shape.ImageData.Save(imageFullPath);
                csvLines.Add($"{imageFileName},{Path.GetFileName(docPath)}");
                imageIndex++;
            }

            if (imageIndex == 0)
                throw new InvalidOperationException($"No images were found in document '{docPath}'.");
        }

        // Write the CSV manifest.
        string manifestPath = Path.Combine(baseDir, "ImageManifest.csv");
        File.WriteAllLines(manifestPath, csvLines);

        // Validate manifest creation.
        if (!File.Exists(manifestPath))
            throw new InvalidOperationException("CSV manifest was not created.");

        // Example completed without interactive prompts.
    }

    private static void CreateSampleDocumentWithImages(string filePath)
    {
        // Simple 1x1 pixel PNG (transparent) encoded in base64.
        const string pngBase64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9yhl4AAAAASUVORK5CYII=";
        byte[] pngBytes = Convert.FromBase64String(pngBase64);

        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert first image.
        using (MemoryStream ms = new MemoryStream(pngBytes))
        {
            builder.InsertImage(ms);
        }

        builder.Writeln(); // Add a paragraph break.

        // Insert second image.
        using (MemoryStream ms = new MemoryStream(pngBytes))
        {
            builder.InsertImage(ms);
        }

        doc.Save(filePath);
    }
}
