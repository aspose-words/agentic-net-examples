using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Define working directories
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string imageDir = Path.Combine(baseDir, "ExtractedImages");
        string sampleImagePath = Path.Combine(baseDir, "sample.png");
        string catalogPath = Path.Combine(baseDir, "Catalog.pdf");

        // Ensure directories exist
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(imageDir);

        // -------------------------------------------------
        // Step 1: Create a deterministic sample image file
        // -------------------------------------------------
        CreateSampleImage(sampleImagePath);

        // -------------------------------------------------
        // Step 2: Generate sample DOCX files containing the image
        // -------------------------------------------------
        int docCount = 3;
        for (int i = 1; i <= docCount; i++)
        {
            string docPath = Path.Combine(inputDir, $"Document{i}.docx");
            CreateSampleDocumentWithImage(docPath, sampleImagePath, $"Sample document {i}");
        }

        // -------------------------------------------------
        // Step 3: Batch process DOCX files, extract images
        // -------------------------------------------------
        List<string> extractedImagePaths = new List<string>();
        foreach (string docFile in Directory.GetFiles(inputDir, "*.docx"))
        {
            Document doc = new Document(docFile);
            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;
            foreach (Shape shape in shapes)
            {
                if (shape.HasImage)
                {
                    string imageFileName = $"{Path.GetFileNameWithoutExtension(docFile)}_Image{imageIndex}.png";
                    string imagePath = Path.Combine(imageDir, imageFileName);
                    shape.ImageData.Save(imagePath);
                    extractedImagePaths.Add(imagePath);
                    imageIndex++;
                }
            }
        }

        // Validate that at least one image was extracted
        if (extractedImagePaths.Count == 0)
            throw new InvalidOperationException("No images were extracted from the DOCX files.");

        // -------------------------------------------------
        // Step 4: Create a PDF catalog summarizing extracted images
        // -------------------------------------------------
        Document catalogDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(catalogDoc);

        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        builder.Font.Size = 16;
        builder.Writeln("Image Catalog");
        builder.Writeln();

        foreach (string imgPath in extractedImagePaths)
        {
            builder.Font.Size = 12;
            builder.Writeln($"Source: {Path.GetFileName(imgPath)}");
            builder.InsertImage(imgPath);
            builder.Writeln();
            builder.InsertBreak(BreakType.PageBreak);
        }

        // Save the catalog as PDF
        catalogDoc.Save(catalogPath, SaveFormat.Pdf);

        // Validate that the PDF catalog was created
        if (!File.Exists(catalogPath))
            throw new InvalidOperationException("Failed to create the PDF catalog.");

        // -------------------------------------------------
        // End of processing
        // -------------------------------------------------
    }

    private static void CreateSampleImage(string filePath)
    {
        // Create a 200x200 white bitmap and draw a simple rectangle
        using (Bitmap bitmap = new Bitmap(200, 200))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.White);
                using (Pen pen = new Pen(Color.Blue, 5))
                {
                    graphics.DrawRectangle(pen, 25, 25, 150, 150);
                }
            }
            // Save as PNG
            bitmap.Save(filePath, ImageFormat.Png);
        }

        // Validate that the image file exists
        if (!File.Exists(filePath))
            throw new InvalidOperationException($"Failed to create sample image at {filePath}");
    }

    private static void CreateSampleDocumentWithImage(string docPath, string imagePath, string title)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Font.Size = 14;
        builder.Writeln(title);
        builder.Writeln();

        // Insert the sample image
        builder.InsertImage(imagePath);
        builder.Writeln();

        // Save the document
        doc.Save(docPath, SaveFormat.Docx);

        // Validate that the document file exists
        if (!File.Exists(docPath))
            throw new InvalidOperationException($"Failed to create sample document at {docPath}");
    }
}
