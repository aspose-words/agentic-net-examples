using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Directories
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string imageOutputDir = Path.Combine(baseDir, "ExtractedImages");
        string reportPath = Path.Combine(baseDir, "ImageReport.csv");

        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(imageOutputDir);

        // Step 1: Create sample images
        string sampleImage1 = Path.Combine(baseDir, "sample1.png");
        string sampleImage2 = Path.Combine(baseDir, "sample2.png");
        CreateSampleImage(sampleImage1, 200, 100, Color.LightBlue, "Img1");
        CreateSampleImage(sampleImage2, 150, 150, Color.LightCoral, "Img2");

        // Step 2: Create sample DOCX files containing the images
        CreateSampleDocx(Path.Combine(inputDir, "DocA.docx"), new[] { sampleImage1, sampleImage2 });
        CreateSampleDocx(Path.Combine(inputDir, "DocB.docx"), new[] { sampleImage2 });

        // Step 3: Batch process DOCX files, extract images, collect metadata
        var rows = new List<string>
        {
            "DocumentName,ImageFileName,Format,SizeBytes,WidthPoints,HeightPoints"
        };

        int extractedCount = 0;

        foreach (string docPath in Directory.GetFiles(inputDir, "*.docx"))
        {
            var doc = new Document(docPath);
            var shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;

            foreach (Shape shape in shapeNodes)
            {
                if (!shape.HasImage) continue;

                string imageFileName = $"{Path.GetFileNameWithoutExtension(docPath)}_Img{imageIndex}.png";
                string imageFullPath = Path.Combine(imageOutputDir, imageFileName);

                // Save image
                shape.ImageData.Save(imageFullPath);

                // Gather metadata
                var fileInfo = new FileInfo(imageFullPath);
                long sizeBytes = fileInfo.Length;
                double widthPoints = shape.ImageData.ImageSize.WidthPoints;
                double heightPoints = shape.ImageData.ImageSize.HeightPoints;
                string format = Path.GetExtension(imageFullPath).TrimStart('.').ToUpperInvariant();

                rows.Add($"{Path.GetFileName(docPath)},{imageFileName},{format},{sizeBytes},{widthPoints},{heightPoints}");
                extractedCount++;
                imageIndex++;
            }
        }

        // Validate that at least one image was extracted
        if (extractedCount == 0)
            throw new InvalidOperationException("No images were extracted from the DOCX files.");

        // Step 4: Write CSV report
        File.WriteAllText(reportPath, string.Join(Environment.NewLine, rows));

        // Validate report file existence
        if (!File.Exists(reportPath))
            throw new InvalidOperationException("CSV report was not created.");
    }

    private static void CreateSampleImage(string path, int width, int height, Color backColor, string text)
    {
        // Create a deterministic bitmap using Aspose.Drawing
        Bitmap bitmap = new Bitmap(width, height);
        Graphics graphics = Graphics.FromImage(bitmap);
        graphics.Clear(backColor);

        // Simple rectangle to have deterministic content
        using (Pen pen = new Pen(Color.Black))
        {
            graphics.DrawRectangle(pen, 0, 0, width - 1, height - 1);
        }

        // Dispose drawing resources and save
        graphics.Dispose();
        bitmap.Save(path);
        bitmap.Dispose();
    }

    private static void CreateSampleDocx(string docPath, string[] imagePaths)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        foreach (string imgPath in imagePaths)
        {
            if (!File.Exists(imgPath))
                throw new FileNotFoundException($"Image file not found: {imgPath}");

            builder.InsertParagraph();
            builder.InsertImage(imgPath);
        }
        doc.Save(docPath);
    }
}
