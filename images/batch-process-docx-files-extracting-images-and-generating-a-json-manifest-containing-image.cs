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
    // Manifest entry model
    public class ImageInfo
    {
        public string Document { get; set; }
        public string ImageFile { get; set; }
        public int Width { get; set; }
        public int Height { get; set; }
    }

    public static void Main()
    {
        // Prepare folders
        string baseDir = Directory.GetCurrentDirectory();
        string inputFolder = Path.Combine(baseDir, "InputDocs");
        string outputImageFolder = Path.Combine(baseDir, "ExtractedImages");
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputImageFolder);

        // Create a deterministic sample image (input.png)
        string sampleImagePath = Path.Combine(baseDir, "input.png");
        CreateSampleImage(sampleImagePath, 200, 150);

        // Create sample DOCX files that contain the image
        CreateSampleDocument(Path.Combine(inputFolder, "Doc1.docx"), sampleImagePath);
        CreateSampleDocument(Path.Combine(inputFolder, "Doc2.docx"), sampleImagePath);

        // Batch process DOCX files
        List<ImageInfo> manifest = new List<ImageInfo>();
        string[] docFiles = Directory.GetFiles(inputFolder, "*.docx");
        foreach (string docPath in docFiles)
        {
            Document doc = new Document(docPath);
            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;
            foreach (Shape shape in shapes)
            {
                if (!shape.HasImage) continue;

                // Save extracted image
                string imageFileName = $"{Path.GetFileNameWithoutExtension(docPath)}_image{imageIndex}.png";
                string imageFullPath = Path.Combine(outputImageFolder, imageFileName);
                shape.ImageData.Save(imageFullPath);

                // Load saved image to obtain dimensions
                using (Aspose.Drawing.Image img = Aspose.Drawing.Image.FromFile(imageFullPath))
                {
                    manifest.Add(new ImageInfo
                    {
                        Document = Path.GetFileName(docPath),
                        ImageFile = imageFileName,
                        Width = img.Width,
                        Height = img.Height
                    });
                }

                imageIndex++;
            }
        }

        // Validation: ensure at least one image was extracted
        if (manifest.Count == 0)
            throw new InvalidOperationException("No images were extracted from the documents.");

        // Serialize manifest to JSON
        string json = JsonConvert.SerializeObject(manifest, Formatting.Indented);
        string manifestPath = Path.Combine(baseDir, "manifest.json");
        File.WriteAllText(manifestPath, json);

        // Validate manifest file creation
        if (!File.Exists(manifestPath))
            throw new InvalidOperationException("Failed to create the JSON manifest file.");
    }

    // Creates a deterministic PNG image using Aspose.Drawing
    private static void CreateSampleImage(string path, int width, int height)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                // Draw a simple rectangle for visual distinction
                using (Pen pen = new Pen(Color.Black, 3))
                {
                    g.DrawRectangle(pen, 10, 10, width - 20, height - 20);
                }
            }
            bitmap.Save(path, ImageFormat.Png);
        }
    }

    // Creates a DOCX file with the specified image inserted
    private static void CreateSampleDocument(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln($"Document: {Path.GetFileName(docPath)}");
        builder.InsertImage(imagePath);
        doc.Save(docPath);
    }
}
