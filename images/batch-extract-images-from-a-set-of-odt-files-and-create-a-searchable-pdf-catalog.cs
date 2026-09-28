using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class BatchImageExtractor
{
    public static void Main()
    {
        // Base working directory
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string imageDir = Path.Combine(baseDir, "ExtractedImages");
        string outputDir = Path.Combine(baseDir, "Output");

        // Ensure required folders exist
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(imageDir);
        Directory.CreateDirectory(outputDir);

        // -------------------------------------------------
        // Step 1: Create a deterministic sample PNG image
        // -------------------------------------------------
        string sampleImagePath = Path.Combine(baseDir, "sample.png");
        CreateSampleImage(sampleImagePath, 200, 150);

        // -------------------------------------------------
        // Step 2: Create sample ODT documents that contain the image
        // -------------------------------------------------
        int docCount = 3;
        for (int i = 1; i <= docCount; i++)
        {
            string docPath = Path.Combine(inputDir, $"Document{i}.odt");
            CreateSampleOdtDocument(docPath, sampleImagePath, $"Sample document {i}");
        }

        // -------------------------------------------------
        // Step 3: Batch extract images from all ODT files
        // -------------------------------------------------
        List<string> extractedImagePaths = new List<string>();
        foreach (string odtFile in Directory.GetFiles(inputDir, "*.odt"))
        {
            Document doc = new Document(odtFile);
            NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;

            foreach (Shape shape in shapeNodes)
            {
                if (shape.HasImage)
                {
                    string ext = GetImageExtension(shape.ImageData.ImageType);
                    string imageFileName = $"{Path.GetFileNameWithoutExtension(odtFile)}_image{imageIndex}{ext}";
                    string imagePath = Path.Combine(imageDir, imageFileName);
                    shape.ImageData.Save(imagePath);
                    extractedImagePaths.Add(imagePath);
                    imageIndex++;
                }
            }
        }

        // Validate that at least one image was extracted
        if (extractedImagePaths.Count == 0)
            throw new InvalidOperationException("No images were extracted from the ODT files.");

        // -------------------------------------------------
        // Step 4: Create a searchable PDF catalog containing the extracted images
        // -------------------------------------------------
        Document catalog = new Document();
        DocumentBuilder builder = new DocumentBuilder(catalog);

        builder.ParagraphFormat.SpaceAfter = 12;
        builder.Font.Size = 14;
        builder.Writeln("Image Catalog");
        builder.Font.Size = 12;
        builder.Writeln($"Generated on {DateTime.Now}");
        builder.InsertParagraph();

        foreach (string imgPath in extractedImagePaths)
        {
            builder.Writeln($"Image from source: {Path.GetFileNameWithoutExtension(imgPath)}");
            builder.InsertImage(imgPath);
            builder.InsertParagraph();
        }

        string catalogPdfPath = Path.Combine(outputDir, "ImageCatalog.pdf");
        catalog.Save(catalogPdfPath, SaveFormat.Pdf);

        // Validate PDF creation
        if (!File.Exists(catalogPdfPath))
            throw new InvalidOperationException("Failed to create the PDF catalog.");
    }

    // Creates a deterministic sample PNG image using Aspose.Drawing
    private static void CreateSampleImage(string filePath, int width, int height)
    {
        // Create bitmap
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height))
        {
            // Obtain graphics object
            using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                // Fill background
                g.Clear(Aspose.Drawing.Color.White);

                // Draw a simple rectangle for visual distinction
                using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Blue, 3))
                {
                    g.DrawRectangle(pen, 10, 10, width - 20, height - 20);
                }
            }

            // Save as PNG
            bitmap.Save(filePath, Aspose.Drawing.Imaging.ImageFormat.Png);
        }
    }

    // Creates a simple ODT document with an image and some text
    private static void CreateSampleOdtDocument(string docPath, string imagePath, string title)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        builder.Font.Size = 16;
        builder.Writeln(title);
        builder.InsertParagraph();

        builder.InsertImage(imagePath);
        builder.InsertParagraph();

        builder.Font.Size = 12;
        builder.Writeln("This document contains a sample image.");

        doc.Save(docPath, SaveFormat.Odt);
    }

    // Maps Aspose.Words ImageType to a file extension
    private static string GetImageExtension(ImageType imageType)
    {
        return imageType switch
        {
            ImageType.Jpeg => ".jpg",
            ImageType.Png => ".png",
            ImageType.Gif => ".gif",
            ImageType.Bmp => ".bmp",
            ImageType.Emf => ".emf",
            ImageType.Wmf => ".wmf",
            _ => ".img"
        };
    }
}
