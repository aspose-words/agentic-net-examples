using System;
using System.IO;
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
        // Define folders.
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string imageDir = Path.Combine(baseDir, "ExtractedImages");
        string reportPath = Path.Combine(baseDir, "ConsolidatedReport.pdf");

        // Ensure clean environment.
        if (Directory.Exists(inputDir)) Directory.Delete(inputDir, true);
        if (Directory.Exists(imageDir)) Directory.Delete(imageDir, true);
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(imageDir);

        // -------------------------------------------------
        // 1. Create a sample image (deterministic file).
        // -------------------------------------------------
        string sampleImagePath = Path.Combine(baseDir, "sample.png");
        CreateSampleImage(sampleImagePath, 200, 200);

        // -------------------------------------------------
        // 2. Create sample DOCX files that contain the image.
        // -------------------------------------------------
        for (int i = 1; i <= 3; i++)
        {
            string docPath = Path.Combine(inputDir, $"Doc{i}.docx");
            CreateDocumentWithImage(docPath, sampleImagePath, $"Document {i}");
        }

        // -------------------------------------------------
        // 3. Batch process: extract images from each DOCX.
        // -------------------------------------------------
        var extractedImagePaths = Directory.GetFiles(inputDir, "*.docx")
            .SelectMany(docFile => ExtractImagesFromDocument(docFile, imageDir))
            .ToList();

        // Validate that at least one image was extracted.
        if (!extractedImagePaths.Any())
            throw new InvalidOperationException("No images were extracted from the input documents.");

        // -------------------------------------------------
        // 4. Build a consolidated PDF report containing all extracted images.
        // -------------------------------------------------
        Document reportDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(reportDoc);

        builder.Writeln("Consolidated Image Report");
        builder.Writeln($"Generated on {DateTime.Now}");
        builder.Writeln();

        foreach (var group in extractedImagePaths.GroupBy(Path.GetFileNameWithoutExtension))
        {
            builder.Writeln($"Images from {group.Key}:");
            foreach (string imgPath in group)
            {
                // Insert the image into the report.
                builder.InsertImage(imgPath);
                builder.Writeln(); // Add a line break after each image.
            }
            builder.InsertBreak(BreakType.PageBreak);
        }

        // Save the report as PDF.
        reportDoc.Save(reportPath, SaveFormat.Pdf);
    }

    // Creates a deterministic PNG image using Aspose.Drawing.
    private static void CreateSampleImage(string filePath, int width, int height)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                // Fill background with white.
                graphics.Clear(Aspose.Drawing.Color.White);
                // Draw a simple rectangle for visual distinction.
                using (Pen pen = new Pen(Aspose.Drawing.Color.Blue, 5))
                {
                    graphics.DrawRectangle(pen, 10, 10, width - 20, height - 20);
                }
            }
            // Save as PNG.
            bitmap.Save(filePath, ImageFormat.Png);
        }
    }

    // Creates a DOCX file with a title and inserts the specified image.
    private static void CreateDocumentWithImage(string docPath, string imagePath, string title)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln(title);
        builder.InsertImage(imagePath);
        builder.Writeln(); // Add a line break.

        doc.Save(docPath);
    }

    // Extracts all images from a DOCX file and saves them to the target folder.
    // Returns the full paths of the saved images.
    private static string[] ExtractImagesFromDocument(string docPath, string targetFolder)
    {
        Document doc = new Document(docPath);
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        var savedImages = shapeNodes
            .OfType<Shape>()
            .Where(s => s.HasImage)
            .Select((shape, index) =>
            {
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string imageFileName = $"{Path.GetFileNameWithoutExtension(docPath)}_Image{index}{extension}";
                string fullPath = Path.Combine(targetFolder, imageFileName);
                shape.ImageData.Save(fullPath);
                return fullPath;
            })
            .ToArray();

        return savedImages;
    }
}
