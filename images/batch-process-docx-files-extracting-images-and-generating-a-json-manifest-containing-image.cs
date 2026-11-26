using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Newtonsoft.Json;

public class Program
{
    // Manifest entry model
    public class ManifestEntry
    {
        public string DocumentName { get; set; }
        public string ImageFileName { get; set; }
        public int Width { get; set; }
        public int Height { get; set; }
    }

    public static void Main()
    {
        // Prepare folders
        string inputFolder = "InputDocs";
        string imageOutputFolder = "ExtractedImages";
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(imageOutputFolder);

        // Create a deterministic sample image (input.png)
        string sampleImagePath = "input.png";
        CreateSampleImage(sampleImagePath, 200, 200);

        // Create sample DOCX files that contain the image
        CreateSampleDocument(Path.Combine(inputFolder, "Doc1.docx"), sampleImagePath);
        CreateSampleDocument(Path.Combine(inputFolder, "Doc2.docx"), sampleImagePath);

        // Batch process DOCX files
        var manifest = new List<ManifestEntry>();
        foreach (string docPath in Directory.GetFiles(inputFolder, "*.docx"))
        {
            var doc = new Document(docPath);
            var shapes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;

            foreach (Shape shape in shapes)
            {
                if (!shape.HasImage) continue;

                string imageFileName = $"{Path.GetFileNameWithoutExtension(docPath)}_Image{imageIndex}.png";
                string imageFullPath = Path.Combine(imageOutputFolder, imageFileName);

                // Save the image from the shape
                shape.ImageData.Save(imageFullPath);

                // Load saved image to obtain dimensions
                using (var img = Image.FromFile(imageFullPath))
                {
                    var entry = new ManifestEntry
                    {
                        DocumentName = Path.GetFileName(docPath),
                        ImageFileName = imageFileName,
                        Width = img.Width,
                        Height = img.Height
                    };
                    manifest.Add(entry);
                }

                // Validate that the image file was created
                if (!File.Exists(imageFullPath))
                    throw new Exception($"Failed to create image file: {imageFullPath}");

                imageIndex++;
            }
        }

        // Validate that at least one image was extracted
        if (manifest.Count == 0)
            throw new Exception("No images were extracted from the documents.");

        // Serialize manifest to JSON
        string manifestJson = JsonConvert.SerializeObject(manifest, Formatting.Indented);
        string manifestPath = "manifest.json";
        File.WriteAllText(manifestPath, manifestJson);

        // Validate manifest file existence
        if (!File.Exists(manifestPath))
            throw new Exception("Failed to create the manifest JSON file.");
    }

    // Creates a deterministic PNG image using Aspose.Drawing
    private static void CreateSampleImage(string path, int width, int height)
    {
        var bitmap = new Bitmap(width, height);
        var graphics = Graphics.FromImage(bitmap);
        graphics.Clear(Color.White);
        // Draw a simple rectangle for visual distinction
        var pen = new Pen(Color.Black, 5);
        graphics.DrawRectangle(pen, 10, 10, width - 20, height - 20);
        pen.Dispose();
        graphics.Dispose();
        bitmap.Save(path, ImageFormat.Png);
        bitmap.Dispose();
    }

    // Creates a DOCX file with the provided image inserted
    private static void CreateSampleDocument(string docPath, string imagePath)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln($"Document: {Path.GetFileName(docPath)}");
        builder.InsertImage(imagePath);
        doc.Save(docPath);
    }
}
