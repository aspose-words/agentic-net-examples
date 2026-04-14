using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Base working directory.
        string baseDir = Directory.GetCurrentDirectory();

        // Directories for input ODT files, extracted images and the manifest.
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string imageDir = Path.Combine(baseDir, "ExtractedImages");
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(imageDir);

        // -----------------------------------------------------------------
        // 1. Create deterministic sample images (must exist before insertion).
        // -----------------------------------------------------------------
        string sampleImage1Path = Path.Combine(baseDir, "sample1.png");
        string sampleImage2Path = Path.Combine(baseDir, "sample2.png");

        CreateSampleImage(sampleImage1Path, 200, 100, Aspose.Drawing.Color.LightBlue);
        CreateSampleImage(sampleImage2Path, 150, 150, Aspose.Drawing.Color.LightCoral);

        // -----------------------------------------------------------------
        // 2. Generate a few ODT documents that contain the sample images.
        // -----------------------------------------------------------------
        for (int docIndex = 1; docIndex <= 2; docIndex++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln($"Sample ODT Document {docIndex}");
            // Insert both sample images.
            builder.InsertImage(sampleImage1Path);
            builder.InsertImage(sampleImage2Path);

            string odtPath = Path.Combine(inputDir, $"Doc{docIndex}.odt");
            doc.Save(odtPath, SaveFormat.Odt);
        }

        // -----------------------------------------------------------------
        // 3. Batch process all ODT files: extract images and build a manifest.
        // -----------------------------------------------------------------
        var manifestEntries = new List<DocumentManifestEntry>();
        string[] odtFiles = Directory.GetFiles(inputDir, "*.odt");

        foreach (string odtFile in odtFiles)
        {
            Document doc = new Document(odtFile);
            NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

            var extractedImages = new List<string>();
            int imageIndex = 0;

            foreach (Shape shape in shapeNodes.OfType<Shape>())
            {
                if (!shape.HasImage)
                    continue; // Skip shapes without images.

                // Determine file extension based on the image type.
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string imageFileName = $"{Path.GetFileNameWithoutExtension(odtFile)}_img{imageIndex}{extension}";
                string imagePath = Path.Combine(imageDir, imageFileName);

                // Save the image to the file system.
                shape.ImageData.Save(imagePath);
                extractedImages.Add(imageFileName);
                imageIndex++;
            }

            // Validation: each document must contain at least one extracted image.
            if (extractedImages.Count == 0)
                throw new InvalidOperationException($"No images were extracted from '{odtFile}'.");

            manifestEntries.Add(new DocumentManifestEntry
            {
                DocumentName = Path.GetFileName(odtFile),
                Images = extractedImages
            });
        }

        // -----------------------------------------------------------------
        // 4. Serialize the manifest to JSON.
        // -----------------------------------------------------------------
        var manifest = new { Documents = manifestEntries };
        string manifestJson = JsonSerializer.Serialize(manifest, new JsonSerializerOptions { WriteIndented = true });
        string manifestPath = Path.Combine(baseDir, "manifest.json");
        File.WriteAllText(manifestPath, manifestJson);

        // Final validation: ensure the manifest file was created.
        if (!File.Exists(manifestPath))
            throw new FileNotFoundException("The JSON manifest was not created.", manifestPath);
    }

    // Helper method to create a deterministic bitmap using Aspose.Drawing.
    private static void CreateSampleImage(string filePath, int width, int height, Aspose.Drawing.Color backgroundColor)
    {
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height);
        Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap);
        graphics.Clear(backgroundColor);
        // Additional deterministic drawing can be added here if needed.
        bitmap.Save(filePath);
        graphics.Dispose();
        bitmap.Dispose();
    }

    // Data structure for the JSON manifest.
    private class DocumentManifestEntry
    {
        public string DocumentName { get; set; }
        public List<string> Images { get; set; }
    }
}
