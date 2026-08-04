using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing; // Aspose.Drawing.Common namespace
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Define folders
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string imageDir = Path.Combine(baseDir, "ExtractedImages");
        string outputDir = Path.Combine(baseDir, "Output");

        // Ensure folders exist
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(imageDir);
        Directory.CreateDirectory(outputDir);

        // Create a deterministic sample image (sample.png)
        string sampleImagePath = Path.Combine(baseDir, "sample.png");
        CreateSampleImage(sampleImagePath, 200, 200);

        // Create sample DOCX files with the sample image inserted
        const int sampleDocCount = 3;
        for (int i = 1; i <= sampleDocCount; i++)
        {
            string docPath = Path.Combine(inputDir, $"Document{i}.docx");
            CreateSampleDocumentWithImage(docPath, sampleImagePath, $"Sample document {i}");
        }

        // List to hold paths of all extracted images
        var extractedImagePaths = new System.Collections.Generic.List<string>();

        // Process each DOCX file in the input folder
        foreach (string docFile in Directory.GetFiles(inputDir, "*.docx"))
        {
            Document doc = new Document(docFile);

            // Get all shape nodes that contain images
            var imageShapes = doc.GetChildNodes(NodeType.Shape, true)
                                 .Cast<Shape>()
                                 .Where(s => s.HasImage)
                                 .ToList();

            if (imageShapes.Count == 0)
                continue; // No images in this document

            int imgIndex = 0;
            foreach (Shape shape in imageShapes)
            {
                // Determine file extension based on image type
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string imageFileName = $"{Path.GetFileNameWithoutExtension(docFile)}_img{imgIndex}{extension}";
                string imageFullPath = Path.Combine(imageDir, imageFileName);

                // Save the image to the file system
                shape.ImageData.Save(imageFullPath);
                extractedImagePaths.Add(imageFullPath);
                imgIndex++;
            }
        }

        // Validate that at least one image was extracted
        if (extractedImagePaths.Count == 0)
            throw new InvalidOperationException("No images were extracted from the input documents.");

        // Create a consolidated PDF report containing all extracted images
        Document reportDoc = new Document();
        DocumentBuilder reportBuilder = new DocumentBuilder(reportDoc);

        reportBuilder.Writeln("Consolidated Image Report");
        reportBuilder.Writeln("---------------------------");
        reportBuilder.Writeln();

        foreach (string imgPath in extractedImagePaths)
        {
            // Add a caption with the image file name
            reportBuilder.Writeln($"Image: {Path.GetFileName(imgPath)}");
            // Insert the image
            reportBuilder.InsertImage(imgPath);
            reportBuilder.Writeln(); // Add spacing between images
        }

        // Save the report as PDF
        string reportPath = Path.Combine(outputDir, "ConsolidatedReport.pdf");
        reportDoc.Save(reportPath, SaveFormat.Pdf);

        Console.WriteLine($"Report generated successfully at: {reportPath}");
    }

    // Creates a simple white bitmap and saves it to the specified path
    private static void CreateSampleImage(string filePath, int width, int height)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.White);
            }
            bitmap.Save(filePath);
        }
    }

    // Creates a DOCX file with a single image and some text
    private static void CreateSampleDocumentWithImage(string docPath, string imagePath, string title)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln(title);
        builder.InsertImage(imagePath);
        builder.Writeln(); // Add a blank line

        doc.Save(docPath);
    }
}
