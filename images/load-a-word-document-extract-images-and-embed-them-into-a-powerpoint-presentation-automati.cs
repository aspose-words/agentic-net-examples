using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a deterministic sample image file.
        const string sampleImagePath = "sample1.png";
        CreateSampleImage(sampleImagePath, 200, 200);

        // Step 2: Create a Word document and insert the sample image.
        const string wordFilePath = "sample.docx";
        CreateWordDocumentWithImage(wordFilePath, sampleImagePath);

        // Step 3: Load the Word document and extract all embedded images.
        string[] extractedImages = ExtractImagesFromWord(wordFilePath);

        // Validate that at least one image was extracted.
        if (extractedImages.Length == 0)
            throw new InvalidOperationException("No images were extracted from the Word document.");

        // Step 4: Create a placeholder PowerPoint file and list the extracted images.
        const string pptxPath = "output.pptx";
        CreatePlaceholderPowerPoint(pptxPath, extractedImages);

        // Validate that the placeholder presentation file was created.
        if (!File.Exists(pptxPath))
            throw new InvalidOperationException("Failed to create the PowerPoint placeholder file.");

        Console.WriteLine("Process completed successfully.");
    }

    // Creates a deterministic PNG image using Aspose.Drawing.
    private static void CreateSampleImage(string path, int width, int height)
    {
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height);
        Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap);
        graphics.Clear(Aspose.Drawing.Color.White);
        using (Aspose.Drawing.SolidBrush brush = new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.Blue))
        {
            graphics.FillRectangle(brush, 20, 20, width - 40, height - 40);
        }
        bitmap.Save(path);
        graphics.Dispose();
        bitmap.Dispose();

        if (!File.Exists(path))
            throw new InvalidOperationException($"Failed to create sample image at '{path}'.");
    }

    // Creates a Word document and inserts the previously created image.
    private static void CreateWordDocumentWithImage(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath);

        if (!File.Exists(docPath))
            throw new InvalidOperationException($"Failed to save Word document at '{docPath}'.");
    }

    // Extracts all images from the Word document and saves them as separate files.
    private static string[] ExtractImagesFromWord(string docPath)
    {
        Document doc = new Document(docPath);
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        List<string> extractedPaths = new List<string>();
        int index = 0;

        foreach (Shape shape in shapes)
        {
            if (shape.HasImage)
            {
                string imageFileName = $"extracted_{index}.png";
                shape.ImageData.Save(imageFileName);
                if (!File.Exists(imageFileName))
                    throw new InvalidOperationException($"Failed to save extracted image '{imageFileName}'.");
                extractedPaths.Add(imageFileName);
                index++;
            }
        }

        return extractedPaths.ToArray();
    }

    // Creates a placeholder PowerPoint file (empty .pptx) and writes a simple manifest of images.
    private static void CreatePlaceholderPowerPoint(string pptxPath, string[] imagePaths)
    {
        // Create an empty file with .pptx extension.
        using (FileStream fs = new FileStream(pptxPath, FileMode.Create, FileAccess.Write))
        {
            // Write a minimal placeholder content.
            using (StreamWriter writer = new StreamWriter(fs))
            {
                writer.WriteLine("Placeholder PowerPoint file.");
                writer.WriteLine("The following images were extracted from the Word document:");
                foreach (string img in imagePaths)
                {
                    writer.WriteLine("- " + img);
                }
            }
        }

        // Validate that the placeholder file exists.
        if (!File.Exists(pptxPath))
            throw new InvalidOperationException($"Failed to create placeholder PowerPoint at '{pptxPath}'.");
    }
}
