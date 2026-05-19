using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Prepare output folder
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // File paths
        string imagePath = Path.Combine(artifactsDir, "input.png");
        string docPath = Path.Combine(artifactsDir, "sample.docx");
        string jsonPath = Path.Combine(artifactsDir, "images.json");

        // 1. Create a deterministic sample image.
        CreateSampleImage(imagePath);

        // 2. Create a Word document that contains the sample image.
        CreateDocumentWithImage(docPath, imagePath);

        // 3. Load the document and extract all images as Base64 strings.
        List<ImageInfo> extractedImages = ExtractImagesToBase64(docPath);

        // 4. Serialize the extracted data to JSON.
        string json = JsonConvert.SerializeObject(extractedImages, Formatting.Indented);
        File.WriteAllText(jsonPath, json);

        // Validation: ensure at least one image was extracted and JSON file exists.
        if (extractedImages.Count == 0 || !File.Exists(jsonPath))
            throw new InvalidOperationException("Image extraction failed or JSON file was not created.");
    }

    // Creates a simple 100x100 white PNG image using Aspose.Drawing.
    private static void CreateSampleImage(string path)
    {
        Bitmap bitmap = new Bitmap(100, 100);
        Graphics graphics = Graphics.FromImage(bitmap);
        graphics.Clear(Color.White);
        bitmap.Save(path);
        graphics.Dispose();
        bitmap.Dispose();
    }

    // Builds a Word document and inserts the image from the given file path.
    private static void CreateDocumentWithImage(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath);
    }

    // Loads a document, iterates over all Shape nodes, extracts image data,
    // converts it to Base64 and records the file extension.
    private static List<ImageInfo> ExtractImagesToBase64(string docPath)
    {
        Document doc = new Document(docPath);
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        var images = new List<ImageInfo>();
        int index = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                ImageData imageData = shape.ImageData;
                byte[] bytes = imageData.ToByteArray();
                string base64 = Convert.ToBase64String(bytes);
                string extension = FileFormatUtil.ImageTypeToExtension(imageData.ImageType);
                images.Add(new ImageInfo
                {
                    Index = index,
                    Extension = extension,
                    Base64 = base64
                });
                index++;
            }
        }

        return images;
    }

    // Simple DTO for JSON serialization.
    private class ImageInfo
    {
        public int Index { get; set; }
        public string Extension { get; set; }
        public string Base64 { get; set; }
    }
}
