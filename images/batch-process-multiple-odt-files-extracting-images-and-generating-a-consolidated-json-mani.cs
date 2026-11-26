using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Define folders for input documents and extracted images.
        string inputFolder = "InputDocs";
        string outputFolder = "ExtractedImages";

        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create a deterministic sample image that will be inserted into the ODT files.
        string sampleImagePath = Path.Combine(inputFolder, "sample.png");
        CreateSampleImage(sampleImagePath, 200, 100);

        // Create a few sample ODT documents, each containing the sample image.
        for (int i = 1; i <= 3; i++)
        {
            string docPath = Path.Combine(inputFolder, $"Doc{i}.odt");
            CreateSampleOdtDocument(docPath, sampleImagePath, $"Sample document {i}");
        }

        // Prepare a manifest structure.
        var manifest = new List<ManifestEntry>();
        int totalExtractedImages = 0;

        // Process each ODT file in the input folder.
        foreach (string odtFile in Directory.GetFiles(inputFolder, "*.odt"))
        {
            var doc = new Document(odtFile);
            var shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
            var extractedImages = new List<string>();
            int imageIndex = 1;

            foreach (Shape shape in shapeNodes)
            {
                if (shape.HasImage)
                {
                    string imageFileName = $"{Path.GetFileNameWithoutExtension(odtFile)}_image{imageIndex}.png";
                    string imagePath = Path.Combine(outputFolder, imageFileName);
                    shape.ImageData.Save(imagePath);
                    extractedImages.Add(imageFileName);
                    imageIndex++;
                }
            }

            totalExtractedImages += extractedImages.Count;

            manifest.Add(new ManifestEntry
            {
                Document = Path.GetFileName(odtFile),
                Images = extractedImages
            });
        }

        // Validate that at least one image was extracted.
        if (totalExtractedImages == 0)
            throw new InvalidOperationException("No images were extracted from the ODT files.");

        // Serialize the manifest to JSON.
        string manifestJson = JsonConvert.SerializeObject(manifest, Formatting.Indented);
        string manifestPath = Path.Combine(outputFolder, "manifest.json");
        File.WriteAllText(manifestPath, manifestJson);

        // Validate that the manifest file was created.
        if (!File.Exists(manifestPath))
            throw new InvalidOperationException("Failed to create the manifest JSON file.");
    }

    // Helper method to create a deterministic PNG image using Aspose.Drawing.
    private static void CreateSampleImage(string path, int width, int height)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.White);
                // Draw a simple rectangle for visual distinction.
                using (var pen = new Pen(Color.Black))
                {
                    graphics.DrawRectangle(pen, 10, 10, width - 20, height - 20);
                }
            }
            bitmap.Save(path, ImageFormat.Png);
        }
    }

    // Helper method to create a simple ODT document with an inserted image.
    private static void CreateSampleOdtDocument(string docPath, string imagePath, string title)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln(title);
        builder.InsertImage(imagePath);
        doc.Save(docPath);
    }

    // Manifest entry structure for JSON serialization.
    private class ManifestEntry
    {
        public string Document { get; set; }
        public List<string> Images { get; set; }
    }
}
