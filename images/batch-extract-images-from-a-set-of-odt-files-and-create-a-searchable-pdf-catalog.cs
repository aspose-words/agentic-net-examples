using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;

public class BatchImageExtractor
{
    public static void Main()
    {
        // Define deterministic folders relative to the executable directory.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string imagesDir = Path.Combine(baseDir, "ExtractedImages");
        string catalogDir = Path.Combine(baseDir, "Catalog");

        // Ensure clean environment.
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(imagesDir);
        Directory.CreateDirectory(catalogDir);

        // Step 1: Create sample images and ODT documents that contain them.
        CreateSampleDocuments(inputDir);

        // Step 2: Batch process ODT files, extract images, and collect info for the catalog.
        var catalogEntries = ProcessDocumentsAndExtractImages(inputDir, imagesDir);

        // Step 3: Build a searchable PDF catalog that lists each source document and its images.
        CreatePdfCatalog(catalogEntries, catalogDir);

        // Validation: ensure at least one image was extracted and catalog PDF exists.
        if (!catalogEntries.Any())
            throw new InvalidOperationException("No images were extracted from the ODT files.");

        string catalogPdfPath = Path.Combine(catalogDir, "ImageCatalog.pdf");
        if (!File.Exists(catalogPdfPath))
            throw new FileNotFoundException("The PDF catalog was not created.", catalogPdfPath);
    }

    // Creates a few ODT files, each containing a deterministic sample image.
    private static void CreateSampleDocuments(string inputDir)
    {
        // Create three sample images.
        for (int i = 0; i < 3; i++)
        {
            string imagePath = Path.Combine(inputDir, $"sample{i}.png");
            CreateSampleImage(imagePath, 200 + i * 50, 150 + i * 30, i);
        }

        // Insert each image into its own ODT document.
        for (int i = 0; i < 3; i++)
        {
            string imagePath = Path.Combine(inputDir, $"sample{i}.png");
            string odtPath = Path.Combine(inputDir, $"Document{i}.odt");

            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"Document {i + 1} containing an image.");
            builder.InsertImage(imagePath);
            doc.Save(odtPath, SaveFormat.Odt);
        }
    }

    // Generates a deterministic PNG image using Aspose.Drawing.
    private static void CreateSampleImage(string filePath, int width, int height, int seed)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            // Fill background with a color derived from the seed.
            int r = (seed * 70) % 256;
            int g = (seed * 130) % 256;
            int b = (seed * 200) % 256;
            graphics.Clear(Color.FromArgb(r, g, b));

            // Draw a simple rectangle border.
            graphics.DrawRectangle(new Pen(Color.White, 3), 5, 5, width - 10, height - 10);

            bitmap.Save(filePath);
        }
    }

    // Processes each ODT file, extracts images, and returns catalog data.
    private static CatalogEntry[] ProcessDocumentsAndExtractImages(string inputDir, string imagesDir)
    {
        var odtFiles = Directory.GetFiles(inputDir, "*.odt");
        var entries = odtFiles.Select(odtPath =>
        {
            var extractedImages = ExtractImagesFromDocument(odtPath, imagesDir);
            return new CatalogEntry
            {
                SourceDocumentName = Path.GetFileName(odtPath),
                ImagePaths = extractedImages
            };
        }).Where(e => e.ImagePaths.Any()).ToArray();

        return entries;
    }

    // Extracts all images from a single document and saves them to the images folder.
    private static string[] ExtractImagesFromDocument(string docPath, string imagesDir)
    {
        Document doc = new Document(docPath);
        var shapeNodes = doc.GetChildNodes(NodeType.Shape, true)
                            .Cast<Shape>()
                            .Where(s => s.HasImage)
                            .ToArray();

        var savedPaths = shapeNodes.Select((shape, index) =>
        {
            string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
            string imageFileName = $"{Path.GetFileNameWithoutExtension(docPath)}_img{index}{extension}";
            string fullPath = Path.Combine(imagesDir, imageFileName);
            shape.ImageData.Save(fullPath);
            return fullPath;
        }).ToArray();

        return savedPaths;
    }

    // Creates a PDF catalog that lists each source document and embeds its extracted images.
    private static void CreatePdfCatalog(CatalogEntry[] entries, string catalogDir)
    {
        Document catalog = new Document();
        DocumentBuilder builder = new DocumentBuilder(catalog);

        // Optional: set PDF save options for better compression.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            ImageCompression = PdfImageCompression.Jpeg,
            JpegQuality = 80
        };

        foreach (var entry in entries)
        {
            // Add a heading for the source document.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
            builder.Writeln(entry.SourceDocumentName);
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;

            // Insert each extracted image.
            foreach (string imgPath in entry.ImagePaths)
            {
                builder.InsertParagraph();
                builder.InsertImage(imgPath);
                builder.InsertParagraph();
                // Add the image file name as searchable text.
                builder.Writeln(Path.GetFileName(imgPath));
            }

            // Add a page break after each document section.
            builder.InsertBreak(BreakType.PageBreak);
        }

        string catalogPath = Path.Combine(catalogDir, "ImageCatalog.pdf");
        catalog.Save(catalogPath, pdfOptions);
    }

    // Simple DTO to hold catalog information.
    private class CatalogEntry
    {
        public string SourceDocumentName { get; set; }
        public string[] ImagePaths { get; set; }
    }
}
