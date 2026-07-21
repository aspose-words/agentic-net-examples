using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Base working directory.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string imageDir = Path.Combine(baseDir, "ExtractedImages");
        string catalogDir = Path.Combine(baseDir, "Catalog");

        // Ensure clean folders.
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(imageDir);
        Directory.CreateDirectory(catalogDir);

        // -----------------------------------------------------------------
        // 1. Create sample images (deterministic local files).
        // -----------------------------------------------------------------
        string sampleImage1 = Path.Combine(baseDir, "sample1.png");
        string sampleImage2 = Path.Combine(baseDir, "sample2.png");

        CreateSampleImage(sampleImage1, 200, 100, Aspose.Drawing.Color.LightBlue, "Img 1");
        CreateSampleImage(sampleImage2, 150, 150, Aspose.Drawing.Color.LightCoral, "Img 2");

        // -----------------------------------------------------------------
        // 2. Create a few ODT documents that contain the sample images.
        // -----------------------------------------------------------------
        string[] odtFiles = new string[3];
        for (int i = 0; i < odtFiles.Length; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln($"Sample ODT document #{i + 1}");
            // Insert both images into each document.
            builder.InsertImage(sampleImage1);
            builder.Writeln(); // line break
            builder.InsertImage(sampleImage2);
            builder.Writeln(); // line break
            builder.Writeln($"End of document #{i + 1}");

            string odtPath = Path.Combine(inputDir, $"SampleDocument{i + 1}.odt");
            doc.Save(odtPath, SaveFormat.Odt);
            odtFiles[i] = odtPath;
        }

        // -----------------------------------------------------------------
        // 3. Batch extract images from all ODT files.
        // -----------------------------------------------------------------
        var catalogBuilder = new DocumentBuilder(new Document());
        int totalExtractedImages = 0;

        foreach (string odtPath in Directory.GetFiles(inputDir, "*.odt"))
        {
            Document srcDoc = new Document(odtPath);
            NodeCollection shapes = srcDoc.GetChildNodes(NodeType.Shape, true);

            // Filter shapes that actually contain images.
            var imageShapes = shapes
                .Cast<Shape>()
                .Where(s => s.HasImage)
                .ToList();

            if (!imageShapes.Any())
                continue; // No images in this document.

            // Add a heading for this source document in the catalog.
            catalogBuilder.Writeln($"Images from: {Path.GetFileName(odtPath)}");
            catalogBuilder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;

            int imgIndex = 0;
            foreach (Shape shape in imageShapes)
            {
                string ext = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string imageFileName = $"{Path.GetFileNameWithoutExtension(odtPath)}_img{imgIndex}{ext}";
                string imagePath = Path.Combine(imageDir, imageFileName);

                // Save the extracted image.
                shape.ImageData.Save(imagePath);
                imgIndex++;
                totalExtractedImages++;

                // Insert the extracted image into the PDF catalog.
                catalogBuilder.InsertImage(imagePath);
                catalogBuilder.Writeln(); // separate images
            }

            // Page break after each source document.
            catalogBuilder.InsertBreak(BreakType.PageBreak);
        }

        // Validate that at least one image was extracted.
        if (totalExtractedImages == 0)
            throw new InvalidOperationException("No images were extracted from the ODT files.");

        // -----------------------------------------------------------------
        // 4. Save the catalog as a searchable PDF.
        // -----------------------------------------------------------------
        string catalogPdfPath = Path.Combine(catalogDir, "ImageCatalog.pdf");
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Use automatic image compression; PDF will be searchable by default.
            ImageCompression = PdfImageCompression.Auto,
            JpegQuality = 80
        };
        catalogBuilder.Document.Save(catalogPdfPath, pdfOptions);

        // Validate that the PDF catalog was created.
        if (!File.Exists(catalogPdfPath))
            throw new FileNotFoundException("Failed to create the PDF catalog.", catalogPdfPath);
    }

    // Helper method to create a deterministic PNG image using Aspose.Drawing.
    private static void CreateSampleImage(string filePath, int width, int height, Aspose.Drawing.Color backColor, string text)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(backColor);
                // Simple text drawing using default font.
                graphics.DrawString(text, new Aspose.Drawing.Font("Arial", 12), Aspose.Drawing.Brushes.Black, new Aspose.Drawing.PointF(10, height / 2 - 10));
            }
            bitmap.Save(filePath);
        }
    }
}
