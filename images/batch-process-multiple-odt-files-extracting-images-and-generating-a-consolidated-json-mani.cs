using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Newtonsoft.Json;

public class Program
{
    public class DocumentManifest
    {
        public string DocumentName { get; set; }
        public List<string> Images { get; set; }
    }

    public static void Main()
    {
        // Base working directory.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        string imagesDir = Path.Combine(baseDir, "Images");
        string docsDir = Path.Combine(baseDir, "Docs");
        string outputDir = Path.Combine(baseDir, "ExtractedImages");
        string manifestPath = Path.Combine(baseDir, "manifest.json");

        // Ensure directories exist.
        Directory.CreateDirectory(imagesDir);
        Directory.CreateDirectory(docsDir);
        Directory.CreateDirectory(outputDir);

        // Create a deterministic sample image.
        string sampleImagePath = Path.Combine(imagesDir, "sample.png");
        CreateSampleImage(sampleImagePath, 200, 200);

        // Create sample ODT documents containing the image.
        const int documentCount = 3;
        for (int i = 1; i <= documentCount; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"Document {i}");
            builder.InsertImage(sampleImagePath);
            string docPath = Path.Combine(docsDir, $"Doc{i}.odt");
            doc.Save(docPath, SaveFormat.Odt);
        }

        // Batch process ODT files.
        var manifest = new List<DocumentManifest>();
        string[] odtFiles = Directory.GetFiles(docsDir, "*.odt");
        foreach (string odtFile in odtFiles)
        {
            Document doc = new Document(odtFile);
            NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
            var extractedImages = new List<string>();
            int imageIndex = 0;

            foreach (Shape shape in shapeNodes.OfType<Shape>())
            {
                if (shape.HasImage)
                {
                    string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                    string imageFileName = $"{Path.GetFileNameWithoutExtension(odtFile)}_img{imageIndex}{extension}";
                    string imageFullPath = Path.Combine(outputDir, imageFileName);
                    shape.ImageData.Save(imageFullPath);
                    extractedImages.Add(imageFileName);
                    imageIndex++;
                }
            }

            if (extractedImages.Count == 0)
                throw new InvalidOperationException($"No images were extracted from '{odtFile}'.");

            manifest.Add(new DocumentManifest
            {
                DocumentName = Path.GetFileName(odtFile),
                Images = extractedImages
            });
        }

        // Serialize manifest to JSON.
        string json = JsonConvert.SerializeObject(manifest, Formatting.Indented);
        File.WriteAllText(manifestPath, json);

        // Validation.
        if (!File.Exists(manifestPath) || manifest.Count == 0)
            throw new InvalidOperationException("Manifest generation failed.");

        Console.WriteLine($"Processed {manifest.Count} document(s). Manifest saved to: {manifestPath}");
    }

    private static void CreateSampleImage(string path, int width, int height)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(Color.LightBlue);
            // Draw a simple rectangle for visual distinction.
            graphics.DrawRectangle(new Pen(Color.DarkBlue, 5), 10, 10, width - 20, height - 20);
            bitmap.Save(path, ImageFormat.Png);
        }
    }
}
