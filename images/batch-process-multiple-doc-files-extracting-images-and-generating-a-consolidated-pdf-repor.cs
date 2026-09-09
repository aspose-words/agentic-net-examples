using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Words.Loading;
using Aspose.Drawing;

public class BatchImageToPdf
{
    public static void Main()
    {
        // Root folder for all temporary data.
        string rootFolder = Path.Combine(Directory.GetCurrentDirectory(), "BatchImageDemo");
        string inputFolder = Path.Combine(rootFolder, "InputDocs");
        string imageFolder = Path.Combine(rootFolder, "ExtractedImages");
        string outputFolder = Path.Combine(rootFolder, "Output");

        // Ensure clean directories.
        foreach (string folder in new[] { inputFolder, imageFolder, outputFolder })
        {
            if (Directory.Exists(folder))
                Directory.Delete(folder, true);
            Directory.CreateDirectory(folder);
        }

        // -------------------------------------------------
        // 1. Create sample images using Aspose.Drawing.
        // -------------------------------------------------
        string sampleImage1 = Path.Combine(rootFolder, "sample1.png");
        string sampleImage2 = Path.Combine(rootFolder, "sample2.png");
        CreateSampleImage(sampleImage1, 200, 200, Aspose.Drawing.Color.Blue);
        CreateSampleImage(sampleImage2, 200, 200, Aspose.Drawing.Color.Green);

        // -------------------------------------------------
        // 2. Create sample DOCX files that contain the images.
        // -------------------------------------------------
        CreateSampleDocument(Path.Combine(inputFolder, "Doc1.docx"), "Document 1", sampleImage1);
        CreateSampleDocument(Path.Combine(inputFolder, "Doc2.docx"), "Document 2", sampleImage2);

        // -------------------------------------------------
        // 3. Batch process each DOCX: extract images.
        // -------------------------------------------------
        var extractedImagesByDoc = new Dictionary<string, List<string>>();

        foreach (string docPath in Directory.GetFiles(inputFolder, "*.docx"))
        {
            var extractedImages = ExtractImagesFromDocument(docPath, imageFolder);
            extractedImagesByDoc[Path.GetFileName(docPath)] = extractedImages;
        }

        // -------------------------------------------------
        // 4. Build a consolidated PDF report that embeds all extracted images.
        // -------------------------------------------------
        Document report = new Document();
        DocumentBuilder builder = new DocumentBuilder(report);

        foreach (var kvp in extractedImagesByDoc)
        {
            string docName = kvp.Key;
            List<string> images = kvp.Value;

            builder.Writeln($"Images extracted from {docName}:");
            builder.Writeln();

            foreach (string imgPath in images)
            {
                // Insert each extracted image into the report.
                builder.InsertImage(imgPath);
                builder.Writeln(); // Add spacing.
            }

            // Separate sections for each source document.
            builder.InsertBreak(BreakType.PageBreak);
        }

        // Save the report as PDF with JPEG compression.
        string pdfPath = Path.Combine(outputFolder, "ConsolidatedReport.pdf");
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            ImageCompression = PdfImageCompression.Jpeg,
            JpegQuality = 80
        };
        report.Save(pdfPath, pdfOptions);

        // Validate that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Failed to create the consolidated PDF report.");

        // Optional: inform that processing completed (no interactive I/O required).
        Console.WriteLine("Processing completed. PDF saved to: " + pdfPath);
    }

    // Creates a simple PNG image with a solid background color.
    private static void CreateSampleImage(string filePath, int width, int height, Aspose.Drawing.Color backColor)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(backColor);
            bitmap.Save(filePath);
        }
    }

    // Creates a DOCX file with a title and a single image.
    private static void CreateSampleDocument(string docPath, string title, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln(title);
        builder.InsertParagraph();
        builder.InsertImage(imagePath);
        builder.InsertParagraph();

        doc.Save(docPath);
    }

    // Extracts all images from a document and saves them to the specified folder.
    // Returns a list of file paths to the saved images.
    private static List<string> ExtractImagesFromDocument(string docPath, string outputFolder)
    {
        List<string> savedImages = new List<string>();
        Document doc = new Document(docPath);

        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string imageFileName = $"{Path.GetFileNameWithoutExtension(docPath)}_Image{imageIndex}{extension}";
                string imageFullPath = Path.Combine(outputFolder, imageFileName);

                shape.ImageData.Save(imageFullPath);
                savedImages.Add(imageFullPath);
                imageIndex++;
            }
        }

        // Ensure at least one image was extracted; otherwise throw.
        if (savedImages.Count == 0)
            throw new InvalidOperationException($"No images found in document '{docPath}'.");

        return savedImages;
    }
}
