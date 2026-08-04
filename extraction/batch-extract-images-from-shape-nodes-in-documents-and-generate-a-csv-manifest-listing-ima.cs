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
        // Prepare folders.
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "Input");
        string outputDir = Path.Combine(baseDir, "Output");
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create sample documents with an embedded image.
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9yhl4AAAAASUVORK5CYII=";
        for (int docIndex = 1; docIndex <= 2; docIndex++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            using (MemoryStream imgStream = new MemoryStream(Convert.FromBase64String(base64Png)))
            {
                builder.InsertImage(imgStream);
            }
            string docPath = Path.Combine(inputDir, $"Sample{docIndex}.docx");
            doc.Save(docPath);
        }

        // Prepare CSV manifest.
        List<string> csvLines = new List<string> { "ImageFile,SourceDocument" };
        int processedDocIndex = 0;

        // Process each document in the input folder.
        foreach (string docFile in Directory.GetFiles(inputDir, "*.docx"))
        {
            processedDocIndex++;
            Document loadedDoc = new Document(docFile);
            var shapes = loadedDoc.GetChildNodes(NodeType.Shape, true).OfType<Shape>();
            int imageIndex = 0;

            foreach (Shape shape in shapes)
            {
                if (!shape.HasImage)
                    continue;

                string imageFileName = $"image-{processedDocIndex}-{imageIndex}{FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType)}";
                string imagePath = Path.Combine(outputDir, imageFileName);
                shape.ImageData.Save(imagePath);
                csvLines.Add($"{imageFileName},{Path.GetFileName(docFile)}");
                imageIndex++;
            }
        }

        // Write the CSV manifest.
        string csvPath = Path.Combine(outputDir, "manifest.csv");
        File.WriteAllLines(csvPath, csvLines);

        // Validation.
        if (!File.Exists(csvPath))
            throw new InvalidOperationException("CSV manifest was not created.");

        if (csvLines.Count <= 1)
            throw new InvalidOperationException("No images were extracted.");

        // Example completed without interactive prompts.
    }
}
