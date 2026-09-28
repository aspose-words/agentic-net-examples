using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Prepare folders
        string baseDir = Directory.GetCurrentDirectory();
        string inputFolder = Path.Combine(baseDir, "InputDocs");
        string outputImageFolder = Path.Combine(baseDir, "ExtractedImages");
        string csvReportPath = Path.Combine(baseDir, "ImageReport.csv");
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputImageFolder);

        // Create a deterministic sample image
        string sampleImagePath = Path.Combine(baseDir, "sample.png");
        CreateSampleImage(sampleImagePath, 100, 100);

        // Create sample DOCX files with images
        CreateSampleDocument(Path.Combine(inputFolder, "Doc1.docx"), sampleImagePath, 2);
        CreateSampleDocument(Path.Combine(inputFolder, "Doc2.docx"), sampleImagePath, 1);

        // Batch process DOCX files
        var csvLines = new List<string>();
        csvLines.Add("DocumentName,ImageFileName,ImageFormat,ImageSizeBytes,ImageWidthPixels,ImageHeightPixels");

        int totalExtractedImages = 0;

        foreach (string docPath in Directory.GetFiles(inputFolder, "*.docx"))
        {
            Document doc = new Document(docPath);
            NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 1;

            foreach (Shape shape in shapeNodes)
            {
                if (!shape.HasImage)
                    continue;

                // Determine output image file name
                string imageExtension = "." + shape.ImageData.ImageType.ToString().ToLower();
                string imageFileName = $"{Path.GetFileNameWithoutExtension(docPath)}_Image{imageIndex}{imageExtension}";
                string imageOutputPath = Path.Combine(outputImageFolder, imageFileName);

                // Save the image
                shape.ImageData.Save(imageOutputPath);
                totalExtractedImages++;

                // Get image metadata
                FileInfo fi = new FileInfo(imageOutputPath);
                long sizeBytes = fi.Length;
                string format = imageExtension.TrimStart('.').ToUpperInvariant();

                // Load image to get dimensions using Aspose.Drawing
                using (Bitmap bmp = new Bitmap(imageOutputPath))
                {
                    int width = bmp.Width;
                    int height = bmp.Height;

                    // Add CSV line
                    string csvLine = $"{Path.GetFileName(docPath)},{imageFileName},{format},{sizeBytes},{width},{height}";
                    csvLines.Add(csvLine);
                }

                imageIndex++;
            }
        }

        // Validate that at least one image was extracted
        if (totalExtractedImages == 0)
        {
            throw new InvalidOperationException("No images were extracted from the DOCX files.");
        }

        // Write CSV report
        File.WriteAllLines(csvReportPath, csvLines, Encoding.UTF8);
    }

    private static void CreateSampleImage(string path, int width, int height)
    {
        // Create bitmap
        Bitmap bitmap = new Bitmap(width, height);
        // Create graphics from bitmap
        Graphics graphics = Graphics.FromImage(bitmap);
        // Fill background
        graphics.Clear(Aspose.Drawing.Color.White);
        // Draw a simple rectangle
        using (Pen pen = new Pen(Aspose.Drawing.Color.Blue, 3))
        {
            graphics.DrawRectangle(pen, 10, 10, width - 20, height - 20);
        }
        // Save image
        bitmap.Save(path, ImageFormat.Png);
        // Clean up
        graphics.Dispose();
        bitmap.Dispose();
    }

    private static void CreateSampleDocument(string docPath, string imagePath, int imageCount)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln($"Sample document: {Path.GetFileName(docPath)}");
        for (int i = 0; i < imageCount; i++)
        {
            builder.InsertImage(imagePath);
            builder.Writeln(); // Add a line break after each image
        }
        doc.Save(docPath);
    }
}
