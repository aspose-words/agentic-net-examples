using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;
using Aspose.Words.Loading;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Newtonsoft.Json;

public class ImageExtractionToMarkdown
{
    public static void Main()
    {
        // Define file names
        const string sampleImagePath = "sample.png";
        const string docPath = "sample.docx";
        const string markdownPath = "output.md";

        // Step 1: Create a deterministic sample image
        CreateSampleImage(sampleImagePath);

        // Step 2: Create a Word document and insert the sample image twice
        CreateWordDocumentWithImages(docPath, sampleImagePath);

        // Step 3: Load the document and extract images
        List<string> extractedImageFiles = ExtractImagesFromDocument(docPath);

        // Validate that at least one image was extracted
        if (extractedImageFiles.Count == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        // Step 4: Generate Markdown file with image links
        GenerateMarkdownFile(markdownPath, extractedImageFiles);

        // Validate that the markdown file was created
        if (!File.Exists(markdownPath))
            throw new FileNotFoundException("Markdown file was not created.", markdownPath);
    }

    private static void CreateSampleImage(string filePath)
    {
        // Create a 100x100 white bitmap
        using (Bitmap bitmap = new Bitmap(100, 100))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.White);
                // Draw a simple black rectangle for visual distinction
                using (Pen pen = new Pen(Color.Black, 2))
                {
                    graphics.DrawRectangle(pen, 10, 10, 80, 80);
                }
            }

            // Save the bitmap as PNG
            bitmap.Save(filePath, ImageFormat.Png);
        }

        // Ensure the image file exists
        if (!File.Exists(filePath))
            throw new FileNotFoundException("Sample image was not created.", filePath);
    }

    private static void CreateWordDocumentWithImages(string docPath, string imagePath)
    {
        // Ensure the source image exists
        if (!File.Exists(imagePath))
            throw new FileNotFoundException("Source image not found.", imagePath);

        // Create a new document and insert the image twice
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert first image
        builder.InsertImage(imagePath);
        builder.Writeln(); // Add a line break

        // Insert second image
        builder.InsertImage(imagePath);
        builder.Writeln();

        // Save the document
        doc.Save(docPath, SaveFormat.Docx);

        // Validate that the document was saved
        if (!File.Exists(docPath))
            throw new FileNotFoundException("Word document was not created.", docPath);
    }

    private static List<string> ExtractImagesFromDocument(string docPath)
    {
        // Load the document
        Document doc = new Document(docPath);

        // Collect extracted image file names
        List<string> imageFiles = new List<string>();

        // Get all Shape nodes in the document
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 1;

        foreach (Shape shape in shapes)
        {
            if (shape.HasImage)
            {
                string imageFileName = $"image-{imageIndex}.png";
                shape.ImageData.Save(imageFileName);
                imageFiles.Add(imageFileName);
                imageIndex++;
            }
        }

        // Validate that each extracted file exists
        foreach (string file in imageFiles)
        {
            if (!File.Exists(file))
                throw new FileNotFoundException("Extracted image file not found.", file);
        }

        return imageFiles;
    }

    private static void GenerateMarkdownFile(string markdownPath, List<string> imageFiles)
    {
        using (StreamWriter writer = new StreamWriter(markdownPath, false))
        {
            foreach (string imageFile in imageFiles)
            {
                writer.WriteLine($"![]({imageFile})");
            }
        }
    }
}
