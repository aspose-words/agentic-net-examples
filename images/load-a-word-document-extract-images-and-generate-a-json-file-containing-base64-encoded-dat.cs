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
    public static void Main()
    {
        // Paths for temporary files
        const string imagePath = "sample.png";
        const string docPath = "sample.docx";
        const string jsonPath = "images.json";

        // -------------------------------------------------
        // 1. Create a deterministic sample image (100x100 white PNG)
        // -------------------------------------------------
        var bitmap = new Bitmap(100, 100);
        var graphics = Graphics.FromImage(bitmap);
        graphics.Clear(Color.White);
        // (Optional) draw a simple rectangle for visual distinction
        graphics.DrawRectangle(new Pen(Color.Black, 2), 10, 10, 80, 80);
        bitmap.Save(imagePath);
        graphics.Dispose();
        bitmap.Dispose();

        // -------------------------------------------------
        // 2. Create a Word document and insert the sample image twice
        // -------------------------------------------------
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        builder.InsertParagraph(); // separate the images
        builder.InsertImage(imagePath);
        doc.Save(docPath);

        // -------------------------------------------------
        // 3. Load the document and extract all images
        // -------------------------------------------------
        var loadedDoc = new Document(docPath);
        var shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        var extractedImages = new List<object>();

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Get raw image bytes
            byte[] imageBytes = shape.ImageData.ToByteArray();

            // Convert to Base64
            string base64 = Convert.ToBase64String(imageBytes);

            // Determine file extension based on image type
            string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);

            // Store information for JSON output
            extractedImages.Add(new
            {
                FileName = $"image{extractedImages.Count}{extension}",
                Base64 = base64,
                ImageType = shape.ImageData.ImageType.ToString()
            });
        }

        // Validate that at least one image was extracted
        if (extractedImages.Count == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        // -------------------------------------------------
        // 4. Serialize the collection to JSON and write to file
        // -------------------------------------------------
        string json = JsonConvert.SerializeObject(extractedImages, Formatting.Indented);
        File.WriteAllText(jsonPath, json);

        // Validate that the JSON file was created
        if (!File.Exists(jsonPath))
            throw new InvalidOperationException("Failed to create the JSON output file.");

        // (Optional) Clean up temporary files – comment out if you need to inspect them
        // File.Delete(imagePath);
        // File.Delete(docPath);
    }
}
