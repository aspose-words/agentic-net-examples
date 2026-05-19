using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Base working directory.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        string odtDir = Path.Combine(baseDir, "OdtFiles");
        string imagesDir = Path.Combine(baseDir, "ExtractedImages");
        string pdfPath = Path.Combine(baseDir, "Catalog.pdf");

        // Ensure clean folders.
        Directory.CreateDirectory(odtDir);
        Directory.CreateDirectory(imagesDir);
        Directory.CreateDirectory(baseDir);

        // -------------------------------------------------
        // 1. Create deterministic sample images.
        // -------------------------------------------------
        string sampleImage1 = Path.Combine(baseDir, "sample1.png");
        string sampleImage2 = Path.Combine(baseDir, "sample2.png");

        CreateSampleImage(sampleImage1, 200, 200, Aspose.Drawing.Color.LightBlue);
        CreateSampleImage(sampleImage2, 200, 200, Aspose.Drawing.Color.LightCoral);

        // -------------------------------------------------
        // 2. Generate a few ODT documents that contain images.
        // -------------------------------------------------
        string[] odtFiles = new string[3];
        for (int i = 0; i < odtFiles.Length; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln($"Document {i + 1}");
            // Insert both sample images into each document.
            builder.InsertImage(sampleImage1);
            builder.Writeln(); // line break
            builder.InsertImage(sampleImage2);
            builder.Writeln(); // line break
            builder.Writeln($"End of Document {i + 1}");

            string odtPath = Path.Combine(odtDir, $"Sample{i + 1}.odt");
            doc.Save(odtPath, SaveFormat.Odt);
            odtFiles[i] = odtPath;
        }

        // -------------------------------------------------
        // 3. Batch extract images from all ODT files.
        // -------------------------------------------------
        List<string> extractedImagePaths = new List<string>();
        int globalImageIndex = 0;

        foreach (string odtFile in Directory.GetFiles(odtDir, "*.odt"))
        {
            Document srcDoc = new Document(odtFile);
            NodeCollection shapeNodes = srcDoc.GetChildNodes(NodeType.Shape, true);

            foreach (Shape shape in shapeNodes.OfType<Shape>())
            {
                if (!shape.HasImage) continue;

                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string imageFileName = $"Img_{globalImageIndex}{extension}";
                string imagePath = Path.Combine(imagesDir, imageFileName);
                shape.ImageData.Save(imagePath);
                extractedImagePaths.Add(imagePath);
                globalImageIndex++;
            }
        }

        // Validate that at least one image was extracted.
        if (extractedImagePaths.Count == 0)
            throw new InvalidOperationException("No images were extracted from the ODT files.");

        // -------------------------------------------------
        // 4. Create a searchable PDF catalog containing the extracted images.
        // -------------------------------------------------
        Document catalog = new Document();
        DocumentBuilder catBuilder = new DocumentBuilder(catalog);

        foreach (string imgPath in extractedImagePaths)
        {
            // Add a caption with the image file name (makes the PDF searchable).
            catBuilder.Writeln(Path.GetFileName(imgPath));
            // Insert the image.
            catBuilder.InsertImage(imgPath);
            // Add a page break after each image for clarity.
            catBuilder.InsertBreak(BreakType.PageBreak);
        }

        // Configure PDF save options (optional compression settings).
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            ImageCompression = PdfImageCompression.Jpeg,
            JpegQuality = 80
        };

        catalog.Save(pdfPath, pdfOptions);

        // Validate that the PDF catalog was created.
        if (!File.Exists(pdfPath))
            throw new FileNotFoundException("Failed to create the PDF catalog.", pdfPath);
    }

    // Helper method to create a deterministic PNG image.
    private static void CreateSampleImage(string filePath, int width, int height, Aspose.Drawing.Color fillColor)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(fillColor);
            bitmap.Save(filePath, ImageFormat.Png);
        }
    }
}
