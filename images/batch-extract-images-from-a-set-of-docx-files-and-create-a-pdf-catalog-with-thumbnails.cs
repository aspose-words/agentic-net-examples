using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class BatchImageExtractorAndPdfCatalog
{
    public static void Main()
    {
        // Base working directory.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string extractedDir = Path.Combine(baseDir, "ExtractedImages");
        string thumbDir = Path.Combine(baseDir, "Thumbnails");
        string outputPdf = Path.Combine(baseDir, "Catalog.pdf");

        // Ensure all directories exist.
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(extractedDir);
        Directory.CreateDirectory(thumbDir);

        // 1. Create deterministic sample images (PNG) using Aspose.Drawing.
        string[] sampleImagePaths = CreateSampleImages(baseDir);

        // 2. Create deterministic sample DOCX files and insert the images.
        CreateSampleDocuments(inputDir, sampleImagePaths);

        // 3. Extract images from each DOCX, save them and create thumbnails.
        var extractedImages = ExtractImagesAndCreateThumbnails(inputDir, extractedDir, thumbDir);

        // 4. Build a PDF catalog that contains the thumbnails.
        BuildPdfCatalog(thumbDir, outputPdf);

        // 5. Validation.
        if (!File.Exists(outputPdf))
            throw new InvalidOperationException("Catalog PDF was not created.");

        if (extractedImages.Count == 0)
            throw new InvalidOperationException("No images were extracted from the documents.");
    }

    // Creates three sample PNG images and returns their file paths.
    private static string[] CreateSampleImages(string baseDir)
    {
        string[] colors = { "Red", "Green", "Blue" };
        var paths = new string[colors.Length];

        for (int i = 0; i < colors.Length; i++)
        {
            string filePath = Path.Combine(baseDir, $"sample{i + 1}.png");
            using (Bitmap bitmap = new Bitmap(200, 200))
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                Aspose.Drawing.Color fillColor = colors[i] switch
                {
                    "Red" => Aspose.Drawing.Color.Red,
                    "Green" => Aspose.Drawing.Color.Green,
                    "Blue" => Aspose.Drawing.Color.Blue,
                    _ => Aspose.Drawing.Color.White
                };
                g.Clear(fillColor);
                bitmap.Save(filePath);
            }
            paths[i] = filePath;
        }

        return paths;
    }

    // Creates two DOCX files, each containing the sample images.
    private static void CreateSampleDocuments(string inputDir, string[] imagePaths)
    {
        for (int docIndex = 1; docIndex <= 2; docIndex++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"Document {docIndex} - Sample content.");

            foreach (string imgPath in imagePaths)
            {
                // Insert each sample image.
                builder.InsertImage(imgPath);
                builder.Writeln(); // Add a line break after each image.
            }

            string docPath = Path.Combine(inputDir, $"Doc{docIndex}.docx");
            doc.Save(docPath);
        }
    }

    // Extracts images from all DOCX files in inputDir, saves them to extractedDir,
    // creates thumbnails in thumbDir, and returns a list of extracted image file paths.
    private static List<string> ExtractImagesAndCreateThumbnails(
        string inputDir,
        string extractedDir,
        string thumbDir)
    {
        var extractedImages = new List<string>();

        // Process each DOCX file.
        foreach (string docPath in Directory.GetFiles(inputDir, "*.docx"))
        {
            Document doc = new Document(docPath);
            NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

            int imageIndex = 0;
            foreach (Shape shape in shapeNodes.OfType<Shape>())
            {
                if (!shape.HasImage)
                    continue;

                // Determine file extension based on image type.
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string imageFileName = $"Extracted_{Path.GetFileNameWithoutExtension(docPath)}_{imageIndex}{extension}";
                string imageFullPath = Path.Combine(extractedDir, imageFileName);

                // Save the original image.
                shape.ImageData.Save(imageFullPath);
                extractedImages.Add(imageFullPath);
                imageIndex++;

                // Create a thumbnail (100x100) using Aspose.Drawing.
                using (Bitmap original = new Bitmap(imageFullPath))
                {
                    const int thumbWidth = 100;
                    const int thumbHeight = 100;
                    using (Bitmap thumb = new Bitmap(thumbWidth, thumbHeight))
                    using (Graphics g = Graphics.FromImage(thumb))
                    {
                        g.Clear(Aspose.Drawing.Color.White);
                        g.DrawImage(original, 0, 0, thumbWidth, thumbHeight);
                        string thumbFileName = Path.GetFileNameWithoutExtension(imageFileName) + "_thumb.png";
                        string thumbFullPath = Path.Combine(thumbDir, thumbFileName);
                        thumb.Save(thumbFullPath);
                    }
                }
            }
        }

        return extractedImages;
    }

    // Builds a PDF document that contains all thumbnail images.
    private static void BuildPdfCatalog(string thumbDir, string outputPdfPath)
    {
        Document catalog = new Document();
        DocumentBuilder builder = new DocumentBuilder(catalog);
        builder.Writeln("Image Catalog");
        builder.Writeln();

        // Insert each thumbnail image.
        foreach (string thumbPath in Directory.GetFiles(thumbDir, "*_thumb.png"))
        {
            builder.InsertImage(thumbPath);
            builder.Writeln(); // Separate images with a line break.
        }

        // Save as PDF.
        catalog.Save(outputPdfPath, SaveFormat.Pdf);
    }
}
