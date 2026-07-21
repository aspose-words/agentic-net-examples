using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing; // Required for Bitmap, Graphics, Color

public class ImageCatalogGenerator
{
    // Creates a deterministic PNG image of the given size and background color.
    private static void CreateSampleImage(string filePath, int width, int height, Aspose.Drawing.Color backgroundColor)
    {
        // Create a bitmap and fill it with the specified background color.
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height);
        Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap);
        graphics.Clear(backgroundColor);

        // Save the bitmap to the supplied file path.
        bitmap.Save(filePath);

        // Clean up drawing resources.
        graphics.Dispose();
        bitmap.Dispose();
    }

    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Set up working folders.
        // -----------------------------------------------------------------
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "ImageCatalogData");
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string imagesDir = Path.Combine(baseDir, "ExtractedImages");
        string catalogPdfPath = Path.Combine(baseDir, "Catalog.pdf");

        Directory.CreateDirectory(baseDir);
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(imagesDir);

        // -----------------------------------------------------------------
        // 2. Create deterministic sample images.
        // -----------------------------------------------------------------
        string sampleImage1 = Path.Combine(baseDir, "sample1.png");
        string sampleImage2 = Path.Combine(baseDir, "sample2.png");

        CreateSampleImage(sampleImage1, 200, 200, Aspose.Drawing.Color.LightBlue);
        CreateSampleImage(sampleImage2, 200, 200, Aspose.Drawing.Color.LightCoral);

        // -----------------------------------------------------------------
        // 3. Generate a few sample DOCX files that embed the images.
        // -----------------------------------------------------------------
        for (int i = 1; i <= 3; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"Sample Document {i}");
            builder.InsertImage(sampleImage1);
            builder.InsertParagraph();
            builder.InsertImage(sampleImage2);
            string docPath = Path.Combine(inputDir, $"Doc{i}.docx");
            doc.Save(docPath);
        }

        // -----------------------------------------------------------------
        // 4. Batch process each DOCX: extract images to the images folder.
        // -----------------------------------------------------------------
        List<string> extractedImagePaths = new List<string>();

        foreach (string docFile in Directory.GetFiles(inputDir, "*.docx"))
        {
            Document doc = new Document(docFile);
            NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

            int imageIndex = 0;
            foreach (Shape shape in shapeNodes.OfType<Shape>())
            {
                if (shape.HasImage)
                {
                    string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                    string imageFileName = $"{Path.GetFileNameWithoutExtension(docFile)}_img{imageIndex}{extension}";
                    string imagePath = Path.Combine(imagesDir, imageFileName);
                    shape.ImageData.Save(imagePath);
                    extractedImagePaths.Add(imagePath);
                    imageIndex++;
                }
            }
        }

        // Validate that at least one image was extracted.
        if (extractedImagePaths.Count == 0)
            throw new InvalidOperationException("No images were extracted from the source documents.");

        // -----------------------------------------------------------------
        // 5. Create a PDF catalog that lists all extracted images.
        // -----------------------------------------------------------------
        Document catalogDoc = new Document();
        DocumentBuilder catalogBuilder = new DocumentBuilder(catalogDoc);
        catalogBuilder.Writeln("Image Catalog");
        catalogBuilder.Font.Size = 16;
        catalogBuilder.Font.Bold = true;
        catalogBuilder.InsertParagraph();

        foreach (string imagePath in extractedImagePaths)
        {
            // Insert the image.
            catalogBuilder.InsertImage(imagePath);
            // Add a caption with the file name.
            catalogBuilder.InsertParagraph();
            catalogBuilder.Writeln(Path.GetFileName(imagePath));
            catalogBuilder.InsertParagraph();
        }

        // Save the catalog as PDF.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            ImageCompression = PdfImageCompression.Auto,
            JpegQuality = 80
        };
        catalogDoc.Save(catalogPdfPath, pdfOptions);

        // -----------------------------------------------------------------
        // 6. Simple validation that the catalog PDF was created.
        // -----------------------------------------------------------------
        if (!File.Exists(catalogPdfPath))
            throw new FileNotFoundException("Failed to create the PDF catalog.", catalogPdfPath);

        Console.WriteLine("Image extraction and PDF catalog generation completed successfully.");
    }
}
