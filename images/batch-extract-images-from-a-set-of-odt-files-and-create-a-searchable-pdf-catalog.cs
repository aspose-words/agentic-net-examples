using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;
using Aspose.Words.Loading;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Root folder for all generated data.
        string rootFolder = Path.Combine(Directory.GetCurrentDirectory(), "BatchImageDemo");
        string odtFolder = Path.Combine(rootFolder, "OdtFiles");
        string extractedImagesFolder = Path.Combine(rootFolder, "ExtractedImages");
        string catalogFolder = Path.Combine(rootFolder, "Catalog");

        // Ensure folders exist.
        Directory.CreateDirectory(odtFolder);
        Directory.CreateDirectory(extractedImagesFolder);
        Directory.CreateDirectory(catalogFolder);

        // 1. Create deterministic sample images.
        string[] sampleImagePaths = CreateSampleImages(rootFolder);

        // 2. Create a few ODT documents that contain those images.
        CreateSampleOdtFiles(odtFolder, sampleImagePaths);

        // 3. Batch extract images from all ODT files.
        List<string> extractedImageFiles = ExtractImagesFromOdtFiles(odtFolder, extractedImagesFolder);

        // Validate that at least one image was extracted.
        if (extractedImageFiles.Count == 0)
            throw new InvalidOperationException("No images were extracted from the ODT files.");

        // 4. Build a searchable PDF catalog that lists the extracted images.
        CreatePdfCatalog(extractedImageFiles, catalogFolder);

        // The example finishes automatically.
    }

    // Creates two simple PNG images using Aspose.Drawing and returns their file paths.
    private static string[] CreateSampleImages(string rootFolder)
    {
        string[] paths = new string[2];
        for (int i = 0; i < 2; i++)
        {
            string filePath = Path.Combine(rootFolder, $"sample{i + 1}.png");
            using (Bitmap bitmap = new Bitmap(200, 200))
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                // Fill with a distinct color.
                Aspose.Drawing.Color fillColor = i == 0
                    ? Aspose.Drawing.Color.FromArgb(255, 100, 150, 200) // Light blue
                    : Aspose.Drawing.Color.FromArgb(255, 200, 150, 100); // Light orange
                g.Clear(fillColor);
                bitmap.Save(filePath);
            }
            paths[i] = filePath;
        }
        return paths;
    }

    // Generates three ODT documents, each containing one of the sample images.
    private static void CreateSampleOdtFiles(string odtFolder, string[] sampleImages)
    {
        for (int i = 0; i < 3; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln($"Document {i + 1} – contains an image.");
            // Alternate between the two sample images.
            string imagePath = sampleImages[i % sampleImages.Length];
            builder.InsertImage(imagePath);

            string odtPath = Path.Combine(odtFolder, $"SampleDocument{i + 1}.odt");
            doc.Save(odtPath, SaveFormat.Odt);
        }
    }

    // Extracts all images from every ODT file in the source folder.
    private static List<string> ExtractImagesFromOdtFiles(string odtFolder, string outputFolder)
    {
        List<string> extractedFiles = new List<string>();
        string[] odtFiles = Directory.GetFiles(odtFolder, "*.odt");

        foreach (string odtPath in odtFiles)
        {
            // Load the ODT document.
            Document doc = new Document(odtPath, new LoadOptions());

            // Get all shape nodes (including images).
            NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;

            foreach (Shape shape in shapeNodes.OfType<Shape>())
            {
                if (!shape.HasImage)
                    continue;

                // Determine proper file extension.
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string imageFileName = $"{Path.GetFileNameWithoutExtension(odtPath)}_img{imageIndex}{extension}";
                string imageFullPath = Path.Combine(outputFolder, imageFileName);

                // Save the image.
                shape.ImageData.Save(imageFullPath);
                extractedFiles.Add(imageFullPath);
                imageIndex++;
            }
        }

        return extractedFiles;
    }

    // Creates a PDF catalog that lists each extracted image with a caption.
    private static void CreatePdfCatalog(List<string> imageFiles, string catalogFolder)
    {
        Document catalog = new Document();
        DocumentBuilder builder = new DocumentBuilder(catalog);

        builder.Writeln("Image Catalog");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln();

        foreach (string imagePath in imageFiles)
        {
            // Insert image.
            builder.InsertImage(imagePath);
            // Add a caption with the file name (searchable text).
            builder.Writeln(Path.GetFileName(imagePath));
            builder.Writeln(); // Add spacing.
        }

        // Configure PDF save options (optional compression).
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            ImageCompression = PdfImageCompression.Jpeg,
            JpegQuality = 80
        };

        string pdfPath = Path.Combine(catalogFolder, "ImageCatalog.pdf");
        catalog.Save(pdfPath, pdfOptions);
    }
}
