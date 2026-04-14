using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class ImageCatalogGenerator
{
    public static void Main()
    {
        // Root folder for all demo data.
        string rootFolder = Path.Combine(Directory.GetCurrentDirectory(), "DemoData");
        string inputFolder = Path.Combine(rootFolder, "InputDocs");
        string imagesFolder = Path.Combine(rootFolder, "ExtractedImages");
        string outputFolder = Path.Combine(rootFolder, "Output");

        // Ensure clean environment.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(imagesFolder);
        Directory.CreateDirectory(outputFolder);

        // 1. Create a deterministic sample image (sample.png).
        string sampleImagePath = Path.Combine(rootFolder, "sample.png");
        CreateSampleImage(sampleImagePath, 200, 200);

        // 2. Generate a few sample DOCX files that contain the sample image.
        const int docCount = 3;
        for (int i = 1; i <= docCount; i++)
        {
            string docPath = Path.Combine(inputFolder, $"Document{i}.docx");
            CreateSampleDocument(docPath, sampleImagePath, $"Sample document {i}");
        }

        // 3. Batch process each DOCX: extract images and collect their paths.
        var extractedImagePaths = new List<string>();
        foreach (string docFile in Directory.GetFiles(inputFolder, "*.docx"))
        {
            var docImages = ExtractImagesFromDocument(docFile, imagesFolder);
            if (docImages.Count == 0)
                throw new InvalidOperationException($"No images were extracted from '{docFile}'.");
            extractedImagePaths.AddRange(docImages);
        }

        // 4. Build a summary PDF catalog that shows all extracted images.
        string catalogPdfPath = Path.Combine(outputFolder, "ImageCatalog.pdf");
        CreatePdfCatalog(catalogPdfPath, extractedImagePaths);

        // 5. Validate final output.
        if (!File.Exists(catalogPdfPath))
            throw new FileNotFoundException("The summary PDF catalog was not created.", catalogPdfPath);

        Console.WriteLine("Processing completed successfully.");
        Console.WriteLine($"Extracted images count: {extractedImagePaths.Count}");
        Console.WriteLine($"Catalog PDF: {catalogPdfPath}");
    }

    // Creates a simple PNG image using Aspose.Drawing.
    private static void CreateSampleImage(string filePath, int width, int height)
    {
        // Ensure any existing file is removed.
        if (File.Exists(filePath))
            File.Delete(filePath);

        Bitmap bitmap = new Bitmap(width, height);
        Graphics graphics = Graphics.FromImage(bitmap);
        graphics.Clear(Color.White);
        // Draw a red ellipse for visual distinction.
        using (Pen pen = new Pen(Color.Red, 5))
        {
            graphics.DrawEllipse(pen, 10, 10, width - 20, height - 20);
        }
        bitmap.Save(filePath);
        graphics.Dispose();
        bitmap.Dispose();
    }

    // Generates a DOCX file with some text and the provided image.
    private static void CreateSampleDocument(string docPath, string imagePath, string title)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln(title);
        builder.InsertParagraph();
        // Insert the deterministic image.
        builder.InsertImage(imagePath);
        doc.Save(docPath);
    }

    // Extracts all images from a DOCX and saves them to the target folder.
    // Returns the full paths of the saved images.
    private static List<string> ExtractImagesFromDocument(string docPath, string targetFolder)
    {
        var savedPaths = new List<string>();
        Document doc = new Document(docPath);
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        int imageIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string imageFileName = $"{Path.GetFileNameWithoutExtension(docPath)}_Image{imageIndex}{extension}";
                string fullPath = Path.Combine(targetFolder, imageFileName);
                shape.ImageData.Save(fullPath);
                savedPaths.Add(fullPath);
                imageIndex++;
            }
        }

        return savedPaths;
    }

    // Creates a PDF that lists all extracted images with captions.
    private static void CreatePdfCatalog(string pdfPath, List<string> imagePaths)
    {
        Document catalog = new Document();
        DocumentBuilder builder = new DocumentBuilder(catalog);
        builder.Writeln("Image Catalog");
        builder.Font.Size = 16;
        builder.Font.Bold = true;
        builder.InsertParagraph();

        foreach (string imgPath in imagePaths)
        {
            // Insert image scaled to a reasonable thumbnail size.
            Shape imgShape = builder.InsertImage(imgPath);
            imgShape.Width = 150; // points
            imgShape.Height = 150;
            builder.Writeln(); // line break after image

            // Caption with file name.
            builder.Writeln(Path.GetFileName(imgPath));
            builder.InsertParagraph();
        }

        // Save as PDF.
        catalog.Save(pdfPath, SaveFormat.Pdf);
    }
}
