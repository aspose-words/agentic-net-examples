using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Words.Loading;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Newtonsoft.Json;

public class Program
{
    // Manifest structures for JSON output
    public class ImageInfo
    {
        public string FileName { get; set; }
        public string ImageType { get; set; }
    }

    public class DocumentInfo
    {
        public string DocumentName { get; set; }
        public List<ImageInfo> Images { get; set; } = new List<ImageInfo>();
    }

    public static void Main()
    {
        // Base folders
        string baseDir = Directory.GetCurrentDirectory();
        string inputFolder = Path.Combine(baseDir, "InputDocs");
        string imageFolder = Path.Combine(baseDir, "ExtractedImages");
        string manifestPath = Path.Combine(baseDir, "Manifest.json");

        // Clean previous run data
        if (Directory.Exists(inputFolder)) Directory.Delete(inputFolder, true);
        if (Directory.Exists(imageFolder)) Directory.Delete(imageFolder, true);
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(imageFolder);

        // Create a deterministic sample image (sample.png)
        string sampleImagePath = Path.Combine(baseDir, "sample.png");
        CreateSampleImage(sampleImagePath, 200, 200);

        // Create a few ODT documents containing the sample image
        CreateSampleOdtDocuments(inputFolder, sampleImagePath, 3);

        // Process each ODT file, extract images, and build manifest
        List<DocumentInfo> manifest = new List<DocumentInfo>();
        foreach (string odtPath in Directory.GetFiles(inputFolder, "*.odt"))
        {
            Document doc = new Document(odtPath);
            NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
            DocumentInfo docInfo = new DocumentInfo { DocumentName = Path.GetFileName(odtPath) };
            int imageIndex = 0;

            foreach (Shape shape in shapeNodes.OfType<Shape>())
            {
                if (!shape.HasImage) continue;

                string extension = Aspose.Words.FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string imageFileName = $"{Path.GetFileNameWithoutExtension(odtPath)}_Image{imageIndex}{extension}";
                string imagePath = Path.Combine(imageFolder, imageFileName);

                shape.ImageData.Save(imagePath);

                docInfo.Images.Add(new ImageInfo
                {
                    FileName = imageFileName,
                    ImageType = shape.ImageData.ImageType.ToString()
                });

                imageIndex++;
            }

            // Validation: ensure at least one image was extracted from this document
            if (docInfo.Images.Count == 0)
                throw new InvalidOperationException($"No images were extracted from document '{odtPath}'.");

            manifest.Add(docInfo);
        }

        // Validation: ensure at least one image was extracted overall
        int totalImages = 0;
        foreach (var d in manifest) totalImages += d.Images.Count;
        if (totalImages == 0)
            throw new InvalidOperationException("No images were extracted from any document.");

        // Serialize manifest to JSON
        string json = JsonConvert.SerializeObject(manifest, Formatting.Indented);
        File.WriteAllText(manifestPath, json);
    }

    // Creates a simple white PNG image with a black rectangle
    private static void CreateSampleImage(string path, int width, int height)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(Aspose.Drawing.Color.White);
            using (Pen pen = new Pen(Aspose.Drawing.Color.Black, 5))
            {
                graphics.DrawRectangle(pen, 10, 10, width - 20, height - 20);
            }
            bitmap.Save(path, ImageFormat.Png);
        }
    }

    // Generates a number of ODT files, each containing the sample image multiple times
    private static void CreateSampleOdtDocuments(string folder, string imagePath, int count)
    {
        for (int i = 1; i <= count; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln($"Document {i}");
            // Insert the sample image twice to ensure multiple images per file
            builder.InsertImage(imagePath);
            builder.InsertParagraph();
            builder.InsertImage(imagePath);

            string odtFile = Path.Combine(folder, $"Doc{i}.odt");
            doc.Save(odtFile, SaveFormat.Odt);
        }
    }
}
