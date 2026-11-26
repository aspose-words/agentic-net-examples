using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Newtonsoft.Json;

public class Program
{
    // Simple record to represent an image entry from a database.
    private class ImageRecord
    {
        public string FilePath { get; set; }
    }

    public static void Main()
    {
        // Prepare deterministic file names.
        const string templatePath = "template.docx";
        const string resultPath = "result.docx";
        const string image1Path = "image1.png";
        const string image2Path = "image2.png";

        // Step 1: Create sample images that will act as database content.
        CreateSampleImage(image1Path, 200, 100, Aspose.Drawing.Color.LightBlue, "Image 1");
        CreateSampleImage(image2Path, 200, 100, Aspose.Drawing.Color.LightGreen, "Image 2");

        // Step 2: Create a simple DOCX template.
        var templateDoc = new Document();
        var templateBuilder = new DocumentBuilder(templateDoc);
        templateBuilder.Writeln("Template Document");
        templateBuilder.Writeln("Images will be inserted below:");
        templateDoc.Save(templatePath);

        // Step 3: Load the template document.
        var doc = new Document(templatePath);
        var builder = new DocumentBuilder(doc);

        // Step 4: Simulate retrieving image records from a database using JSON.
        string json = $"[{{\"FilePath\":\"{image1Path}\"}},{{\"FilePath\":\"{image2Path}\"}}]";
        var imageRecords = JsonConvert.DeserializeObject<List<ImageRecord>>(json);

        // Step 5: Insert each image into the document.
        foreach (var record in imageRecords)
        {
            if (!File.Exists(record.FilePath))
                throw new FileNotFoundException($"Image file not found: {record.FilePath}");

            // Move to a new paragraph for each image.
            builder.Writeln();
            builder.InsertImage(record.FilePath);
        }

        // Step 6: Save the resulting document.
        doc.Save(resultPath);

        // Validation: ensure the output file was created.
        if (!File.Exists(resultPath))
            throw new Exception("Result document was not created.");

        // Optional: verify that at least one image shape exists in the document.
        int imageCount = 0;
        foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
        {
            if (shape.HasImage && shape.ImageData.ImageType != ImageType.Unknown)
                imageCount++;
        }

        if (imageCount == 0)
            throw new Exception("No images were inserted into the document.");

        // Clean up created files (optional, comment out if you want to keep them).
        //File.Delete(templatePath);
        //File.Delete(resultPath);
        //File.Delete(image1Path);
        //File.Delete(image2Path);
    }

    // Helper method to create a deterministic PNG image.
    private static void CreateSampleImage(string filePath, int width, int height, Aspose.Drawing.Color backgroundColor, string text)
    {
        using (var bitmap = new Bitmap(width, height, PixelFormat.Format32bppArgb))
        {
            using (var graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(backgroundColor);
                // Draw simple text if needed (optional).
                // Note: Font usage is avoided to keep the example simple and avoid ambiguity.
            }

            // Save the bitmap as PNG.
            bitmap.Save(filePath, ImageFormat.Png);
        }
    }
}
