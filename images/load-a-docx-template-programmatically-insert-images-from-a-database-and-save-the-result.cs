using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Create sample images that will act as database entries.
        CreateSampleImage("image1.png", 200, 150, Aspose.Drawing.Color.LightBlue);
        CreateSampleImage("image2.png", 200, 150, Aspose.Drawing.Color.LightGreen);

        // Simulate a database storing image file names in JSON format.
        string jsonDatabase = "[\"image1.png\", \"image2.png\"]";
        List<string> imagePaths = JsonConvert.DeserializeObject<List<string>>(jsonDatabase);

        // Create a simple DOCX template.
        const string templatePath = "template.docx";
        CreateTemplateDocument(templatePath);

        // Load the template.
        Document doc = new Document(templatePath);
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert each image from the "database" into the document.
        foreach (string imagePath in imagePaths)
        {
            if (!File.Exists(imagePath))
                throw new FileNotFoundException($"Image file not found: {imagePath}");

            builder.InsertImage(imagePath);
            builder.Writeln(); // Add a line break after each image.
        }

        // Save the resulting document.
        const string outputPath = "result.docx";
        doc.Save(outputPath);

        // Validate that the output file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create output document: {outputPath}");
    }

    private static void CreateSampleImage(string fileName, int width, int height, Aspose.Drawing.Color fillColor)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(fillColor);
            }
            bitmap.Save(fileName, ImageFormat.Png);
        }

        // Ensure the image file exists.
        if (!File.Exists(fileName))
            throw new InvalidOperationException($"Failed to create sample image: {fileName}");
    }

    private static void CreateTemplateDocument(string filePath)
    {
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("This is a template document.");
        builder.Writeln("Images will be inserted below:");
        templateDoc.Save(filePath);

        // Ensure the template file exists.
        if (!File.Exists(filePath))
            throw new InvalidOperationException($"Failed to create template document: {filePath}");
    }
}
