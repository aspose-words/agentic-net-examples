using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Loading;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Base working directory
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "BatchProcessing");
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string outputDir = Path.Combine(baseDir, "ExtractedImages");
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create sample images
        string imagePath1 = Path.Combine(baseDir, "sample1.png");
        CreateSampleImage(imagePath1, 100, 100, Aspose.Drawing.Color.LightBlue);
        string imagePath2 = Path.Combine(baseDir, "sample2.png");
        CreateSampleImage(imagePath2, 120, 80, Aspose.Drawing.Color.LightCoral);

        // Create sample ODT documents containing the images
        CreateSampleDocument(Path.Combine(inputDir, "Doc1.odt"), imagePath1);
        CreateSampleDocument(Path.Combine(inputDir, "Doc2.odt"), imagePath2);
        CreateSampleDocument(Path.Combine(inputDir, "Doc3.odt"), imagePath1, imagePath2); // document with two images

        // Prepare manifest collection
        var manifest = new List<ManifestEntry>();
        int globalImageIndex = 1;

        // Process each ODT file
        foreach (string docPath in Directory.GetFiles(inputDir, "*.odt"))
        {
            var loadOptions = new LoadOptions { LoadFormat = LoadFormat.Odt };
            Document doc = new Document(docPath, loadOptions);

            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
            foreach (Shape shape in shapes)
            {
                if (shape.HasImage)
                {
                    string imageFileName = $"{Path.GetFileNameWithoutExtension(docPath)}_img{globalImageIndex}.png";
                    string imageFullPath = Path.Combine(outputDir, imageFileName);
                    shape.ImageData.Save(imageFullPath);

                    // Record manifest entry
                    manifest.Add(new ManifestEntry
                    {
                        SourceDocument = Path.GetFileName(docPath),
                        ImageFile = imageFileName,
                        ImageIndex = globalImageIndex
                    });

                    globalImageIndex++;
                }
            }
        }

        // Validate that images were extracted
        if (manifest.Count == 0)
            throw new InvalidOperationException("No images were extracted from the ODT files.");

        // Serialize manifest to JSON
        string manifestJson = JsonConvert.SerializeObject(manifest, Formatting.Indented);
        string manifestPath = Path.Combine(baseDir, "manifest.json");
        File.WriteAllText(manifestPath, manifestJson);

        // Validate manifest file creation
        if (!File.Exists(manifestPath))
            throw new InvalidOperationException("Failed to create the manifest JSON file.");
    }

    // Creates a deterministic PNG image using Aspose.Drawing
    private static void CreateSampleImage(string filePath, int width, int height, Aspose.Drawing.Color fillColor)
    {
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height);
        Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap);
        graphics.Clear(fillColor);
        bitmap.Save(filePath);
        graphics.Dispose();
        bitmap.Dispose();
    }

    // Creates an ODT document and inserts the provided images
    private static void CreateSampleDocument(string docPath, params string[] imagePaths)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        foreach (string imgPath in imagePaths)
        {
            if (!File.Exists(imgPath))
                throw new FileNotFoundException($"Image file not found: {imgPath}");
            builder.InsertImage(imgPath);
            builder.Writeln(); // separate images
        }
        doc.Save(docPath, SaveFormat.Odt);
    }

    // Manifest entry definition
    private class ManifestEntry
    {
        public string SourceDocument { get; set; }
        public string ImageFile { get; set; }
        public int ImageIndex { get; set; }
    }
}
