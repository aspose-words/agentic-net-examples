using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;
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
        string outputFolder = Path.Combine(baseDir, "Output");
        string extractedImagesFolder = Path.Combine(outputFolder, "ExtractedImages");

        // Ensure folders exist
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);
        Directory.CreateDirectory(extractedImagesFolder);

        // Create a deterministic sample image (input.png)
        string sampleImagePath = Path.Combine(baseDir, "input.png");
        CreateSampleImage(sampleImagePath, 200, 200, Aspose.Drawing.Color.LightBlue, "Sample");

        // Create sample DOCX files with the image inserted
        int docCount = 2;
        for (int i = 1; i <= docCount; i++)
        {
            string docPath = Path.Combine(inputFolder, $"Doc{i}.docx");
            CreateSampleDocumentWithImage(docPath, sampleImagePath, $"Document {i}");
        }

        // List to hold paths of all extracted images
        List<string> extractedImagePaths = new List<string>();

        // Process each DOCX file: extract images
        foreach (string docFile in Directory.GetFiles(inputFolder, "*.docx"))
        {
            Document doc = new Document(docFile);
            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;
            foreach (Shape shape in shapes)
            {
                if (shape.HasImage)
                {
                    string imageFileName = $"{Path.GetFileNameWithoutExtension(docFile)}_image{imageIndex}.png";
                    string imagePath = Path.Combine(extractedImagesFolder, imageFileName);
                    shape.ImageData.Save(imagePath);
                    extractedImagePaths.Add(imagePath);
                    imageIndex++;
                }
            }
        }

        // Validate that at least one image was extracted
        if (extractedImagePaths.Count == 0)
            throw new InvalidOperationException("No images were extracted from the DOC files.");

        // Create a consolidated PDF report containing all extracted images
        Document pdfReport = new Document();
        DocumentBuilder builder = new DocumentBuilder(pdfReport);

        for (int i = 0; i < extractedImagePaths.Count; i++)
        {
            if (i > 0)
                builder.InsertBreak(BreakType.PageBreak);

            builder.Writeln($"Image {i + 1}");
            builder.InsertImage(extractedImagePaths[i]);
        }

        string pdfPath = Path.Combine(outputFolder, "Report.pdf");
        pdfReport.Save(pdfPath, SaveFormat.Pdf);

        // Validate PDF creation
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Failed to create the consolidated PDF report.");

        // Optional: write simple console output (no waiting)
        Console.WriteLine($"Extracted {extractedImagePaths.Count} images and generated PDF report at: {pdfPath}");
    }

    // Creates a deterministic PNG image with solid background and optional text
    private static void CreateSampleImage(string filePath, int width, int height, Aspose.Drawing.Color backColor, string text)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(backColor);
                // Simple text drawing (optional)
                using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 20))
                {
                    graphics.DrawString(text, font, new SolidBrush(Aspose.Drawing.Color.Black), new PointF(10, height / 2 - 10));
                }
            }
            bitmap.Save(filePath, ImageFormat.Png);
        }
    }

    // Creates a DOCX file with a paragraph and inserts the provided image
    private static void CreateSampleDocumentWithImage(string docPath, string imagePath, string title)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln(title);
        builder.InsertParagraph();
        builder.InsertImage(imagePath);
        doc.Save(docPath);
    }
}
