using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Words.Loading;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Base working directory.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string extractedImagesDir = Path.Combine(baseDir, "ExtractedImages");
        string summaryPdfPath = Path.Combine(baseDir, "SummaryCatalog.pdf");

        // Ensure clean folders.
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(extractedImagesDir);
        if (File.Exists(summaryPdfPath))
            File.Delete(summaryPdfPath);

        // -------------------------------------------------
        // 1. Create deterministic sample images (PNG).
        // -------------------------------------------------
        string sampleImage1Path = Path.Combine(baseDir, "sample1.png");
        string sampleImage2Path = Path.Combine(baseDir, "sample2.png");

        CreateSamplePng(sampleImage1Path, 200, 150, Aspose.Drawing.Color.LightBlue);
        CreateSamplePng(sampleImage2Path, 150, 200, Aspose.Drawing.Color.LightCoral);

        // -------------------------------------------------
        // 2. Create sample DOCX files that contain the images.
        // -------------------------------------------------
        for (int i = 1; i <= 3; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln($"Document {i}");
            // Insert both sample images.
            builder.InsertImage(sampleImage1Path);
            builder.InsertParagraph();
            builder.InsertImage(sampleImage2Path);
            builder.InsertParagraph();

            string docPath = Path.Combine(inputDir, $"Doc{i}.docx");
            doc.Save(docPath);
        }

        // -------------------------------------------------
        // 3. Batch process each DOCX: extract images.
        // -------------------------------------------------
        var extractedImagePaths = new System.Collections.Generic.List<string>();

        foreach (string docFile in Directory.GetFiles(inputDir, "*.docx"))
        {
            Document doc = new Document(docFile);
            var shapeNodes = doc.GetChildNodes(NodeType.Shape, true)
                                .Cast<Shape>()
                                .Where(s => s.HasImage)
                                .ToList();

            int imageIndex = 0;
            foreach (Shape shape in shapeNodes)
            {
                // Determine file extension based on image type.
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string imageFileName = $"{Path.GetFileNameWithoutExtension(docFile)}_Image{imageIndex}{extension}";
                string imageFullPath = Path.Combine(extractedImagesDir, imageFileName);

                // Save the image to the file system.
                shape.ImageData.Save(imageFullPath);
                extractedImagePaths.Add(imageFullPath);
                imageIndex++;
            }
        }

        // Validate that at least one image was extracted.
        if (extractedImagePaths.Count == 0)
            throw new InvalidOperationException("No images were extracted from the documents.");

        // -------------------------------------------------
        // 4. Create a summary PDF catalog containing the extracted images.
        // -------------------------------------------------
        Document summaryDoc = new Document();
        DocumentBuilder summaryBuilder = new DocumentBuilder(summaryDoc);

        summaryBuilder.Writeln("Summary Image Catalog");
        summaryBuilder.Writeln("----------------------");
        summaryBuilder.Writeln();

        foreach (string imagePath in extractedImagePaths)
        {
            // Add a caption with the image file name.
            summaryBuilder.Writeln($"Image: {Path.GetFileName(imagePath)}");
            // Insert the image into the PDF.
            summaryBuilder.InsertImage(imagePath);
            summaryBuilder.Writeln(); // Space between images.
        }

        // Save as PDF with default options (you can customize PdfSaveOptions if needed).
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Example: compress images using JPEG compression.
            ImageCompression = PdfImageCompression.Jpeg,
            JpegQuality = 80
        };
        summaryDoc.Save(summaryPdfPath, pdfOptions);

        // Validate that the summary PDF was created.
        if (!File.Exists(summaryPdfPath))
            throw new InvalidOperationException("Failed to create the summary PDF catalog.");

        // The program finishes without waiting for user input.
    }

    // Helper method to create a deterministic PNG image using Aspose.Drawing.
    private static void CreateSamplePng(string filePath, int width, int height, Aspose.Drawing.Color backgroundColor)
    {
        // Create bitmap.
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            // Create graphics object.
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                // Fill background.
                graphics.Clear(backgroundColor);
            }

            // Save bitmap to PNG file.
            bitmap.Save(filePath);
        }
    }
}
