using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Loading;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Define folders.
        string baseDir = Directory.GetCurrentDirectory();
        string inputFolder = Path.Combine(baseDir, "InputDocs");
        string imageFolder = Path.Combine(baseDir, "ExtractedImages");
        string manifestPath = Path.Combine(baseDir, "Manifest.json");

        // Ensure clean environment.
        if (Directory.Exists(inputFolder)) Directory.Delete(inputFolder, true);
        if (Directory.Exists(imageFolder)) Directory.Delete(imageFolder, true);
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(imageFolder);

        // Create a deterministic sample image (sample.png).
        string sampleImagePath = Path.Combine(baseDir, "sample.png");
        CreateSampleImage(sampleImagePath, 100, 100);

        // Create a few ODT documents containing the sample image.
        const int documentCount = 3;
        for (int i = 1; i <= documentCount; i++)
        {
            string docPath = Path.Combine(inputFolder, $"Doc{i}.odt");
            CreateOdtWithImages(docPath, sampleImagePath, i);
        }

        // Batch process ODT files: extract images and build manifest.
        var manifest = new List<ManifestEntry>();
        var odtFiles = Directory.GetFiles(inputFolder, "*.odt");
        int totalExtractedImages = 0;

        foreach (string odtFile in odtFiles)
        {
            var entry = new ManifestEntry
            {
                Document = Path.GetFileName(odtFile),
                Images = new List<string>()
            };

            // Load the document.
            var loadOptions = new LoadOptions(); // default options
            Document doc = new Document(odtFile, loadOptions);

            // Find all shapes that contain images.
            var shapeNodes = doc.GetChildNodes(NodeType.Shape, true)
                                .Cast<Shape>()
                                .Where(s => s.HasImage)
                                .ToList();

            int imageIndex = 0;
            foreach (Shape shape in shapeNodes)
            {
                // Determine file extension based on image type.
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string imageFileName = $"{Path.GetFileNameWithoutExtension(odtFile)}_Image{imageIndex}{extension}";
                string imageFullPath = Path.Combine(imageFolder, imageFileName);

                // Save the image.
                shape.ImageData.Save(imageFullPath);
                entry.Images.Add(imageFileName);
                imageIndex++;
                totalExtractedImages++;
            }

            manifest.Add(entry);
        }

        // Validate that at least one image was extracted.
        if (totalExtractedImages == 0)
            throw new InvalidOperationException("No images were extracted from the ODT files.");

        // Serialize manifest to JSON.
        string json = JsonConvert.SerializeObject(manifest, Formatting.Indented);
        File.WriteAllText(manifestPath, json);

        // Simple verification output (optional, not interactive).
        Console.WriteLine($"Processed {odtFiles.Length} ODT files.");
        Console.WriteLine($"Extracted {totalExtractedImages} images to '{imageFolder}'.");
        Console.WriteLine($"Manifest written to '{manifestPath}'.");
    }

    // Creates a simple white PNG image using Aspose.Drawing.
    private static void CreateSampleImage(string filePath, int width, int height)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Aspose.Drawing.Color.White);
            }
            bitmap.Save(filePath, ImageFormat.Png);
        }
    }

    // Creates an ODT document with a given number of inserted images.
    private static void CreateOdtWithImages(string docPath, string imagePath, int imageCount)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        for (int i = 0; i < imageCount; i++)
        {
            builder.Writeln($"Image {i + 1} in {Path.GetFileName(docPath)}:");
            builder.InsertImage(imagePath);
            builder.InsertBreak(BreakType.PageBreak);
        }

        // Save as ODT.
        doc.Save(docPath, SaveFormat.Odt);
    }

    // Helper class for JSON manifest entries.
    private class ManifestEntry
    {
        public string Document { get; set; }
        public List<string> Images { get; set; }
    }
}
