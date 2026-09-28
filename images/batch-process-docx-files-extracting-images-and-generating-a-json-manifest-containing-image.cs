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
    // Entry point
    public static void Main()
    {
        // Prepare deterministic folders
        string baseDir = Directory.GetCurrentDirectory();
        string inputFolder = Path.Combine(baseDir, "InputDocs");
        string imageOutputFolder = Path.Combine(baseDir, "ExtractedImages");
        string manifestPath = Path.Combine(baseDir, "manifest.json");

        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(imageOutputFolder);

        // Create a sample image that will be inserted into documents
        string sampleImagePath = Path.Combine(baseDir, "sample.png");
        CreateSampleImage(sampleImagePath, 200, 150);

        // Create sample DOCX files containing the image
        CreateSampleDocument(Path.Combine(inputFolder, "Doc1.docx"), sampleImagePath);
        CreateSampleDocument(Path.Combine(inputFolder, "Doc2.docx"), sampleImagePath);

        // Process each DOCX file: extract images and build manifest entries
        var manifestEntries = new List<ManifestEntry>();
        foreach (string docPath in Directory.GetFiles(inputFolder, "*.docx"))
        {
            Document doc = new Document(docPath);
            NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;

            foreach (Shape shape in shapeNodes)
            {
                if (!shape.HasImage) continue;

                // Save the extracted image
                string imageFileName = $"{Path.GetFileNameWithoutExtension(docPath)}_image{imageIndex}.png";
                string imageFilePath = Path.Combine(imageOutputFolder, imageFileName);
                shape.ImageData.Save(imageFilePath);

                // Load image to obtain pixel dimensions
                using (MemoryStream ms = new MemoryStream(shape.ImageData.ImageBytes))
                {
                    ms.Position = 0;
                    using (Bitmap bmp = new Bitmap(ms))
                    {
                        var entry = new ManifestEntry
                        {
                            DocumentName = Path.GetFileName(docPath),
                            ImageFileName = imageFileName,
                            WidthPixels = bmp.Width,
                            HeightPixels = bmp.Height
                        };
                        manifestEntries.Add(entry);
                    }
                }

                imageIndex++;
            }
        }

        // Validation: ensure at least one image was extracted
        if (manifestEntries.Count == 0)
            throw new InvalidOperationException("No images were extracted from the DOCX files.");

        // Serialize manifest to JSON
        string json = JsonConvert.SerializeObject(manifestEntries, Formatting.Indented);
        File.WriteAllText(manifestPath, json);

        // Final validation: ensure manifest file exists
        if (!File.Exists(manifestPath))
            throw new InvalidOperationException("Failed to create the manifest JSON file.");
    }

    // Creates a deterministic sample PNG image
    private static void CreateSampleImage(string path, int width, int height)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Aspose.Drawing.Color.White);
                // Draw a simple rectangle for visual distinction
                using (Pen pen = new Pen(Aspose.Drawing.Color.Blue, 5))
                {
                    g.DrawRectangle(pen, 10, 10, width - 20, height - 20);
                }
            }
            bitmap.Save(path, ImageFormat.Png);
        }
    }

    // Creates a DOCX file with the provided image inserted
    private static void CreateSampleDocument(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln($"Sample document: {Path.GetFileName(docPath)}");
        builder.InsertImage(imagePath);
        doc.Save(docPath);
    }

    // Manifest entry definition
    private class ManifestEntry
    {
        public string DocumentName { get; set; }
        public string ImageFileName { get; set; }
        public int WidthPixels { get; set; }
        public int HeightPixels { get; set; }
    }
}
