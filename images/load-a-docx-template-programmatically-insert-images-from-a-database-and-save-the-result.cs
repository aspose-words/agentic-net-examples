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
    // Simple model to simulate a database record that contains an image file path.
    private class ImageRecord
    {
        public string FilePath { get; set; }
    }

    public static void Main()
    {
        // Step 1: Create deterministic sample images that will act as "database" images.
        string[] sampleImagePaths = { "image1.png", "image2.png" };
        CreateSampleImage(sampleImagePaths[0], 200, 100, Aspose.Drawing.Color.LightBlue, "Img 1");
        CreateSampleImage(sampleImagePaths[1], 200, 100, Aspose.Drawing.Color.LightGreen, "Img 2");

        // Step 2: Simulate storing image records in a database as JSON.
        var dbRecords = new List<ImageRecord>
        {
            new ImageRecord { FilePath = sampleImagePaths[0] },
            new ImageRecord { FilePath = sampleImagePaths[1] }
        };
        string jsonDb = JsonConvert.SerializeObject(dbRecords);
        // Simulate retrieving the records from the "database".
        var retrievedRecords = JsonConvert.DeserializeObject<List<ImageRecord>>(jsonDb);

        // Step 3: Create a simple DOCX template file.
        const string templatePath = "Template.docx";
        CreateTemplateDocument(templatePath);

        // Step 4: Load the DOCX template.
        Document doc = new Document(templatePath);
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Step 5: Insert each image from the simulated database into the document.
        foreach (var record in retrievedRecords)
        {
            if (!File.Exists(record.FilePath))
                throw new FileNotFoundException($"Image file not found: {record.FilePath}");

            // Insert a line break before each image for readability.
            builder.Writeln();
            builder.InsertImage(record.FilePath);
        }

        // Step 6: Save the resulting document.
        const string resultPath = "Result.docx";
        doc.Save(resultPath);

        // Step 7: Validate that the output file was created.
        if (!File.Exists(resultPath))
            throw new InvalidOperationException($"Failed to create output document: {resultPath}");
    }

    // Helper method to create a deterministic PNG image using Aspose.Drawing.
    private static void CreateSampleImage(string filePath, int width, int height, Aspose.Drawing.Color backgroundColor, string text)
    {
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height))
        {
            using (Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                graphics.Clear(backgroundColor);
                // Draw simple text in the center.
                using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 16))
                {
                    var textSize = graphics.MeasureString(text, font);
                    float x = (width - textSize.Width) / 2;
                    float y = (height - textSize.Height) / 2;
                    graphics.DrawString(text, font, new SolidBrush(Aspose.Drawing.Color.Black), x, y);
                }
            }
            // Save the bitmap as PNG.
            bitmap.Save(filePath, Aspose.Drawing.Imaging.ImageFormat.Png);
        }
    }

    // Helper method to create a minimal DOCX template.
    private static void CreateTemplateDocument(string filePath)
    {
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("This is a template document.");
        builder.Writeln("Images will be inserted below:");
        templateDoc.Save(filePath);
    }
}
