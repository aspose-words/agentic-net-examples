using System;
using System.IO;
using System.Text;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

public class Program
{
    // Entry point of the console application.
    public static void Main()
    {
        // Define folders for the sample document, extracted images and the CSV manifest.
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputFolder);

        string sampleDocPath = Path.Combine(outputFolder, "Sample.docx");
        string imagesFolder = Path.Combine(outputFolder, "Images");
        Directory.CreateDirectory(imagesFolder);
        string manifestPath = Path.Combine(outputFolder, "manifest.csv");

        // Create a sample document that contains a few images.
        CreateSampleDocument(sampleDocPath);

        // Extract images from the document and generate the CSV manifest.
        ExtractImagesAndCreateManifest(sampleDocPath, imagesFolder, manifestPath);

        // Simple confirmation (no interactive input).
        Console.WriteLine("Extraction completed. Manifest saved to: " + manifestPath);
    }

    // Creates a DOCX file with a few embedded images.
    private static void CreateSampleDocument(string docPath)
    {
        // Small 1x1 pixel PNG (transparent) encoded in base64.
        const string pngBase64 =
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+BAQAE/wJ" +
            "9V6cAAAAASUVORK5CYII=";

        byte[] pngBytes = Convert.FromBase64String(pngBase64);

        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert first image.
        using (MemoryStream ms = new MemoryStream(pngBytes))
        {
            Shape imgShape1 = builder.InsertImage(ms);
            imgShape1.Name = "FirstImage";
        }

        builder.Writeln(); // Add a line break between images.

        // Insert second image.
        using (MemoryStream ms = new MemoryStream(pngBytes))
        {
            Shape imgShape2 = builder.InsertImage(ms);
            imgShape2.Name = "SecondImage";
        }

        // Save the document.
        doc.Save(docPath);
    }

    // Extracts all images from shape nodes and writes a CSV manifest.
    private static void ExtractImagesAndCreateManifest(string docPath, string imagesFolder, string csvPath)
    {
        // Load the document.
        Document doc = new Document(docPath);

        // Get all shape nodes (including nested) from the document.
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        // Filter shapes that actually contain an image.
        var imageShapes = shapeNodes
            .OfType<Shape>()
            .Where(s => s.HasImage)
            .ToList();

        if (!imageShapes.Any())
        {
            throw new InvalidOperationException("No image-bearing shapes were found in the document.");
        }

        var csvBuilder = new StringBuilder();
        csvBuilder.AppendLine("ImageFileName,ShapeName");

        int imageIndex = 0;
        foreach (Shape shape in imageShapes)
        {
            // Determine file extension based on the image type.
            string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
            string imageFileName = $"Image_{imageIndex}{extension}";
            string imageFullPath = Path.Combine(imagesFolder, imageFileName);

            // Save the image to the file system.
            shape.ImageData.Save(imageFullPath);

            // Record entry in CSV.
            string shapeName = string.IsNullOrEmpty(shape.Name) ? "UnnamedShape" : shape.Name;
            csvBuilder.AppendLine($"{imageFileName},{shapeName}");

            imageIndex++;
        }

        // Write the CSV manifest.
        File.WriteAllText(csvPath, csvBuilder.ToString(), Encoding.UTF8);

        // Validate that the manifest file was created.
        if (!File.Exists(csvPath))
        {
            throw new IOException("Failed to create the CSV manifest file.");
        }
    }
}
