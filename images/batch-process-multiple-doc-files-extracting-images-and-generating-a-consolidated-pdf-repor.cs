using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Define folders
        string baseDir = Directory.GetCurrentDirectory();
        string inputFolder = Path.Combine(baseDir, "InputDocs");
        string extractedFolder = Path.Combine(baseDir, "ExtractedImages");
        string outputFolder = Path.Combine(baseDir, "Output");

        // Ensure folders exist
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(extractedFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample images that will be inserted into the DOC files
        string sampleImagePath1 = Path.Combine(baseDir, "sample1.png");
        string sampleImagePath2 = Path.Combine(baseDir, "sample2.png");
        CreateSampleImage(sampleImagePath1, 200, 150, Aspose.Drawing.Color.LightBlue, "Img1");
        CreateSampleImage(sampleImagePath2, 250, 180, Aspose.Drawing.Color.LightGreen, "Img2");

        // Create sample DOCX files containing the images
        CreateSampleDocument(Path.Combine(inputFolder, "Doc1.docx"), new[] { sampleImagePath1, sampleImagePath2 });
        CreateSampleDocument(Path.Combine(inputFolder, "Doc2.docx"), new[] { sampleImagePath2 });

        // List to hold paths of all extracted images
        List<string> extractedImagePaths = new List<string>();

        // Process each DOCX file in the input folder
        foreach (string docPath in Directory.GetFiles(inputFolder, "*.docx"))
        {
            Document doc = new Document(docPath);
            NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;

            foreach (Shape shape in shapeNodes)
            {
                if (shape.HasImage)
                {
                    string imageFileName = $"{Path.GetFileNameWithoutExtension(docPath)}_Image{imageIndex}.png";
                    string imageFullPath = Path.Combine(extractedFolder, imageFileName);
                    shape.ImageData.Save(imageFullPath);
                    extractedImagePaths.Add(imageFullPath);
                    imageIndex++;
                }
            }
        }

        // Validate that at least one image was extracted
        if (extractedImagePaths.Count == 0)
            throw new InvalidOperationException("No images were extracted from the DOC files.");

        // Create a consolidated PDF report containing all extracted images
        Document reportDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(reportDoc);

        for (int i = 0; i < extractedImagePaths.Count; i++)
        {
            string imgPath = extractedImagePaths[i];
            // Insert image
            builder.InsertImage(imgPath);
            // Add a page break after each image except the last one
            if (i < extractedImagePaths.Count - 1)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Save the PDF report
        string pdfReportPath = Path.Combine(outputFolder, "ConsolidatedReport.pdf");
        reportDoc.Save(pdfReportPath, SaveFormat.Pdf);

        // Validate that the PDF was created
        if (!File.Exists(pdfReportPath))
            throw new InvalidOperationException("Failed to create the consolidated PDF report.");
    }

    // Helper method to create a deterministic sample PNG image
    private static void CreateSampleImage(string filePath, int width, int height, Aspose.Drawing.Color backColor, string text)
    {
        // Create bitmap
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height))
        {
            // Create graphics from bitmap
            using (Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                // Fill background
                graphics.Clear(backColor);
                // Draw simple text to identify the image
                using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 20))
                {
                    graphics.DrawString(text, font, new SolidBrush(Aspose.Drawing.Color.Black), new PointF(10, height / 2 - 10));
                }
                // No explicit Flush needed; disposing will commit drawing operations
            }
            // Save bitmap as PNG
            bitmap.Save(filePath, ImageFormat.Png);
        }
    }

    // Helper method to create a DOCX file and insert provided images
    private static void CreateSampleDocument(string docPath, string[] imagePaths)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        foreach (string imgPath in imagePaths)
        {
            // Ensure the image file exists before insertion
            if (!File.Exists(imgPath))
                throw new FileNotFoundException($"Image file not found: {imgPath}");
            builder.InsertImage(imgPath);
            builder.Writeln(); // Add a line break between images
        }
        doc.Save(docPath);
    }
}
