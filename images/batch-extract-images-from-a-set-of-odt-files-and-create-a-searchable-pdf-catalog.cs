using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Define folders and files
        string inputFolder = "InputDocs";
        string imageFolder = "ExtractedImages";
        string sampleImagePath = "sample.png";
        string catalogPdfPath = "Catalog.pdf";

        // Ensure clean environment
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(imageFolder);
        if (File.Exists(sampleImagePath)) File.Delete(sampleImagePath);
        if (File.Exists(catalogPdfPath)) File.Delete(catalogPdfPath);
        foreach (var file in Directory.GetFiles(inputFolder, "*.odt"))
            File.Delete(file);
        foreach (var file in Directory.GetFiles(imageFolder, "*.*"))
            File.Delete(file);

        // Create a deterministic sample image
        CreateSampleImage(sampleImagePath);

        // Create sample ODT documents containing the image
        CreateSampleOdtDocuments(inputFolder, sampleImagePath, 3);

        // Batch extract images from ODT files
        List<string> extractedImagePaths = ExtractImagesFromOdtFiles(inputFolder, imageFolder);

        // Validate that images were extracted
        if (extractedImagePaths.Count == 0)
            throw new InvalidOperationException("No images were extracted from the ODT files.");

        // Create a searchable PDF catalog containing the extracted images
        CreatePdfCatalog(extractedImagePaths, catalogPdfPath);

        // Validate that the PDF catalog was created
        if (!File.Exists(catalogPdfPath))
            throw new InvalidOperationException("The PDF catalog was not created.");
    }

    // Creates a simple deterministic PNG image using Aspose.Drawing.
    private static void CreateSampleImage(string path)
    {
        const int width = 200;
        const int height = 200;

        // Create bitmap and graphics objects from Aspose.Drawing.
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                // Draw a simple rectangle for visual distinction.
                using (var pen = new Pen(Color.Black, 3))
                {
                    g.DrawRectangle(pen, 10, 10, width - 20, height - 20);
                }
            }
            // Save the bitmap as PNG.
            bitmap.Save(path);
        }
    }

    // Generates a number of ODT documents that each contain the sample image.
    private static void CreateSampleOdtDocuments(string folder, string imagePath, int count)
    {
        for (int i = 1; i <= count; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"Sample ODT Document {i}");
            builder.InsertImage(imagePath);
            builder.Writeln($"End of document {i}.");
            string odtPath = Path.Combine(folder, $"SampleDocument{i}.odt");
            doc.Save(odtPath, SaveFormat.Odt);
        }
    }

    // Extracts all images from ODT files in the input folder and saves them to the output folder.
    private static List<string> ExtractImagesFromOdtFiles(string inputFolder, string outputFolder)
    {
        List<string> extractedPaths = new List<string>();
        string[] odtFiles = Directory.GetFiles(inputFolder, "*.odt");

        foreach (string odtFile in odtFiles)
        {
            Document doc = new Document(odtFile);
            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 1;

            foreach (Shape shape in shapes)
            {
                if (shape.HasImage)
                {
                    string imageFileName = $"{Path.GetFileNameWithoutExtension(odtFile)}_Image{imageIndex}{GetImageExtension(shape.ImageData.ImageType)}";
                    string imagePath = Path.Combine(outputFolder, imageFileName);
                    shape.ImageData.Save(imagePath);
                    extractedPaths.Add(imagePath);
                    imageIndex++;
                }
            }
        }

        return extractedPaths;
    }

    // Maps Aspose.Words.ImageType values to common file extensions.
    private static string GetImageExtension(ImageType imageType)
    {
        switch (imageType)
        {
            case ImageType.Jpeg: return ".jpg";
            case ImageType.Png:  return ".png";
            case ImageType.Gif:  return ".gif";
            case ImageType.Bmp:  return ".bmp";
            case ImageType.Emf:  return ".emf";
            case ImageType.Wmf:  return ".wmf";
            default:             return ".img";
        }
    }

    // Creates a searchable PDF catalog that lists each extracted image and embeds the image itself.
    private static void CreatePdfCatalog(List<string> imagePaths, string pdfPath)
    {
        Document catalog = new Document();
        DocumentBuilder builder = new DocumentBuilder(catalog);

        // Title
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Image Catalog");
        builder.Writeln(); // Empty line after title

        foreach (string imgPath in imagePaths)
        {
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln($"Image: {Path.GetFileName(imgPath)}");
            builder.InsertImage(imgPath);
            builder.Writeln(); // Spacing between entries
        }

        catalog.Save(pdfPath, SaveFormat.Pdf);
    }
}
