using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;               // Aspose.Drawing.Common namespace
using Aspose.Drawing.Imaging;
using Newtonsoft.Json;

public class Program
{
    // Model for JSON manifest
    public class DocumentManifest
    {
        public string DocumentName { get; set; }
        public List<string> Images { get; set; } = new List<string>();
    }

    public static void Main()
    {
        // Define folders
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string imagesDir = Path.Combine(baseDir, "ExtractedImages");
        string manifestPath = Path.Combine(baseDir, "manifest.json");

        // Ensure clean environment
        if (Directory.Exists(inputDir)) Directory.Delete(inputDir, true);
        if (Directory.Exists(imagesDir)) Directory.Delete(imagesDir, true);
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(imagesDir);

        // -------------------------------------------------
        // Step 1: Create deterministic sample images
        // -------------------------------------------------
        string sampleImagePath = Path.Combine(baseDir, "sample.png");
        CreateSampleImage(sampleImagePath, 200, 200, Aspose.Drawing.Color.LightBlue);

        // -------------------------------------------------
        // Step 2: Create sample ODT documents containing the image
        // -------------------------------------------------
        for (int docIndex = 1; docIndex <= 2; docIndex++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a paragraph and the sample image
            builder.Writeln($"Document {docIndex} with an image.");
            builder.InsertImage(sampleImagePath);

            string odtPath = Path.Combine(inputDir, $"SampleDocument{docIndex}.odt");
            doc.Save(odtPath, SaveFormat.Odt);
        }

        // -------------------------------------------------
        // Step 3: Batch process ODT files, extract images, build manifest
        // -------------------------------------------------
        List<DocumentManifest> manifest = new List<DocumentManifest>();

        foreach (string odtFile in Directory.GetFiles(inputDir, "*.odt"))
        {
            Document doc = new Document(odtFile);
            string docName = Path.GetFileName(odtFile);

            // Collect shapes that actually contain images
            var imageShapes = doc.GetChildNodes(NodeType.Shape, true)
                                 .Cast<Shape>()
                                 .Where(s => s.HasImage)
                                 .ToList();

            if (!imageShapes.Any())
                throw new InvalidOperationException($"No images found in document '{docName}'.");

            DocumentManifest entry = new DocumentManifest { DocumentName = docName };

            int imageIndex = 0;
            foreach (Shape shape in imageShapes)
            {
                // Determine proper file extension for the image type
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string imageFileName = $"{Path.GetFileNameWithoutExtension(docName)}_image_{imageIndex}{extension}";
                string imageFullPath = Path.Combine(imagesDir, imageFileName);

                // Save the image to disk
                shape.ImageData.Save(imageFullPath);
                entry.Images.Add(imageFileName);
                imageIndex++;
            }

            manifest.Add(entry);
        }

        // -------------------------------------------------
        // Step 4: Serialize manifest to JSON
        // -------------------------------------------------
        string json = JsonConvert.SerializeObject(manifest, Formatting.Indented);
        File.WriteAllText(manifestPath, json);

        // Simple validation output
        Console.WriteLine($"Processed {manifest.Count} document(s).");
        Console.WriteLine($"Extracted images are stored in: {imagesDir}");
        Console.WriteLine($"JSON manifest written to: {manifestPath}");
    }

    // Helper: creates a deterministic bitmap and saves it to a file
    private static void CreateSampleImage(string filePath, int width, int height, Aspose.Drawing.Color background)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(background);
            // Draw a simple rectangle for visual distinction
            using (Pen pen = new Pen(Aspose.Drawing.Color.DarkBlue, 5))
            {
                graphics.DrawRectangle(pen, 10, 10, width - 20, height - 20);
            }
            bitmap.Save(filePath, ImageFormat.Png);
        }
    }
}
