using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Define paths
        string baseDir = Directory.GetCurrentDirectory();
        string inputFolder = Path.Combine(baseDir, "InputDocs");
        string extractedFolder = Path.Combine(baseDir, "ExtractedImages");
        string sampleImagePath = Path.Combine(baseDir, "sample.png");
        string reportPath = Path.Combine(baseDir, "Report.pdf");

        // Ensure folders exist
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(extractedFolder);

        // -------------------------------------------------
        // Step 1: Create a deterministic sample image
        // -------------------------------------------------
        const int imgWidth = 200;
        const int imgHeight = 200;
        using (Bitmap bitmap = new Bitmap(imgWidth, imgHeight))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Aspose.Drawing.Color.White);
                // Draw a simple rectangle for visual distinction
                using (Pen pen = new Pen(Aspose.Drawing.Color.Blue, 5))
                {
                    g.DrawRectangle(pen, 10, 10, imgWidth - 20, imgHeight - 20);
                }
            }
            bitmap.Save(sampleImagePath, ImageFormat.Png);
        }

        // -------------------------------------------------
        // Step 2: Create sample DOCX files that contain the image
        // -------------------------------------------------
        for (int i = 1; i <= 2; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"Document {i}");
            builder.InsertParagraph();
            builder.InsertImage(sampleImagePath);
            string docPath = Path.Combine(inputFolder, $"doc{i}.docx");
            doc.Save(docPath);
        }

        // -------------------------------------------------
        // Step 3: Batch process DOCX files, extract images
        // -------------------------------------------------
        List<string> extractedImagePaths = new List<string>();
        string[] docFiles = Directory.GetFiles(inputFolder, "*.docx");
        foreach (string docFile in docFiles)
        {
            Document doc = new Document(docFile);
            NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;
            foreach (Shape shape in shapeNodes)
            {
                if (shape.HasImage)
                {
                    string imageFileName = $"{Path.GetFileNameWithoutExtension(docFile)}_{imageIndex}.png";
                    string imagePath = Path.Combine(extractedFolder, imageFileName);
                    shape.ImageData.Save(imagePath);
                    extractedImagePaths.Add(imagePath);
                    imageIndex++;
                }
            }
        }

        // Validate that at least one image was extracted
        if (extractedImagePaths.Count == 0)
        {
            throw new InvalidOperationException("No images were extracted from the input documents.");
        }

        // -------------------------------------------------
        // Step 4: Generate a consolidated PDF report with extracted images
        // -------------------------------------------------
        Document reportDoc = new Document();
        DocumentBuilder reportBuilder = new DocumentBuilder(reportDoc);
        foreach (string imgPath in extractedImagePaths)
        {
            reportBuilder.Writeln($"Image extracted from: {Path.GetFileName(imgPath)}");
            reportBuilder.InsertParagraph();
            reportBuilder.InsertImage(imgPath);
            reportBuilder.InsertBreak(BreakType.PageBreak);
        }

        // Save the report as PDF
        reportDoc.Save(reportPath, SaveFormat.Pdf);

        // Validate that the PDF report was created
        if (!File.Exists(reportPath))
        {
            throw new InvalidOperationException("Failed to create the PDF report.");
        }

        // Cleanup: (optional) delete temporary files if desired
        // File.Delete(sampleImagePath);
        // foreach (var file in Directory.GetFiles(inputFolder)) File.Delete(file);
        // foreach (var file in Directory.GetFiles(extractedFolder)) File.Delete(file);
    }
}
