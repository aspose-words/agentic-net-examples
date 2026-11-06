using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Newtonsoft.Json;

public class Program
{
    // Represents a single entry in the JSON manifest.
    public class ManifestEntry
    {
        public string Document { get; set; }
        public string ImageFile { get; set; }
        public int ImageIndex { get; set; }
    }

    public static void Main()
    {
        // Define folders used by the example.
        string inputFolder = "InputDocs";
        string imageOutputFolder = "ExtractedImages";
        string manifestPath = "manifest.json";

        // Ensure the required directories exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(imageOutputFolder);

        // Create deterministic sample images that will be inserted into the ODT files.
        string sampleImage1Path = "sample1.png";
        string sampleImage2Path = "sample2.png";

        CreateSampleImage(sampleImage1Path, 100, 100, Aspose.Drawing.Color.LightBlue);
        CreateSampleImage(sampleImage2Path, 120, 80, Aspose.Drawing.Color.LightCoral);

        // Create two sample ODT documents, each containing one of the images above.
        CreateSampleDocument(Path.Combine(inputFolder, "doc1.odt"), sampleImage1Path);
        CreateSampleDocument(Path.Combine(inputFolder, "doc2.odt"), sampleImage2Path);

        // Collection that will become the JSON manifest.
        List<ManifestEntry> manifest = new List<ManifestEntry>();

        // Process every ODT file in the input folder.
        foreach (string docPath in Directory.GetFiles(inputFolder, "*.odt"))
        {
            // Load the document.
            Document doc = new Document(docPath);

            // Get all Shape nodes (images are stored in Shape objects).
            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIdx = 0;

            foreach (Shape shape in shapes)
            {
                // Skip non‑image shapes.
                if (!shape.HasImage)
                    continue;

                imageIdx++;
                string imageFileName = $"{Path.GetFileNameWithoutExtension(docPath)}_image{imageIdx}.png";
                string imageFullPath = Path.Combine(imageOutputFolder, imageFileName);

                // Save the extracted image to disk.
                shape.ImageData.Save(imageFullPath);

                // Verify that the image file was created.
                if (!File.Exists(imageFullPath))
                    throw new InvalidOperationException($"Failed to save extracted image: {imageFullPath}");

                // Add an entry to the manifest.
                manifest.Add(new ManifestEntry
                {
                    Document = Path.GetFileName(docPath),
                    ImageFile = imageFileName,
                    ImageIndex = imageIdx
                });
            }
        }

        // Ensure that at least one image was extracted.
        if (manifest.Count == 0)
            throw new InvalidOperationException("No images were extracted from the ODT files.");

        // Serialize the manifest to a nicely formatted JSON file.
        string json = JsonConvert.SerializeObject(manifest, Formatting.Indented);
        File.WriteAllText(manifestPath, json);

        // Verify that the manifest file was written.
        if (!File.Exists(manifestPath))
            throw new InvalidOperationException("Failed to write the manifest JSON file.");
    }

    // Creates a deterministic PNG image using Aspose.Drawing.
    private static void CreateSampleImage(string filePath, int width, int height, Aspose.Drawing.Color backColor)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(backColor);
            }
            bitmap.Save(filePath);
        }

        // Verify that the image file exists.
        if (!File.Exists(filePath))
            throw new InvalidOperationException($"Failed to create sample image: {filePath}");
    }

    // Creates a simple ODT document that contains the specified image.
    private static void CreateSampleDocument(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln($"Document: {Path.GetFileName(docPath)}");
        builder.InsertImage(imagePath);
        doc.Save(docPath, SaveFormat.Odt);

        // Verify that the document file exists.
        if (!File.Exists(docPath))
            throw new InvalidOperationException($"Failed to create sample document: {docPath}");
    }
}
