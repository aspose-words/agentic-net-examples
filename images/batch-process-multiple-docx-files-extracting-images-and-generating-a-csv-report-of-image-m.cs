using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    // Entry point of the console application.
    public static void Main()
    {
        // Base working directory.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string imageOutputDir = Path.Combine(baseDir, "ExtractedImages");
        string csvReportPath = Path.Combine(baseDir, "ImageReport.csv");

        // Ensure all directories exist.
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(imageOutputDir);

        // -----------------------------------------------------------------
        // 1. Create deterministic sample images (PNG) that will be inserted.
        // -----------------------------------------------------------------
        string sampleImagePath1 = Path.Combine(baseDir, "sample1.png");
        string sampleImagePath2 = Path.Combine(baseDir, "sample2.png");

        CreateSamplePng(sampleImagePath1, 200, 200, Aspose.Drawing.Color.Red);
        CreateSamplePng(sampleImagePath2, 150, 150, Aspose.Drawing.Color.Blue);

        // ---------------------------------------------------------------
        // 2. Create a few sample DOCX files that contain the images above.
        // ---------------------------------------------------------------
        for (int docIndex = 1; docIndex <= 3; docIndex++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert both sample images into each document.
            builder.InsertParagraph();
            builder.Writeln($"Document {docIndex} - first image:");
            builder.InsertImage(sampleImagePath1);

            builder.InsertParagraph();
            builder.Writeln($"Document {docIndex} - second image:");
            builder.InsertImage(sampleImagePath2);

            string docPath = Path.Combine(inputDir, $"SampleDoc{docIndex}.docx");
            doc.Save(docPath);
        }

        // ---------------------------------------------------------------
        // 3. Batch process all DOCX files: extract images and collect metadata.
        // ---------------------------------------------------------------
        var csvLines = new List<string>();
        // Header for CSV.
        csvLines.Add("DocumentName,ImageFileName,ImageType,WidthPoints,HeightPoints,WidthPixels,HeightPixels,HorizontalResolutionDPI,VerticalResolutionDPI");

        bool anyImageExtracted = false;

        foreach (string docFilePath in Directory.GetFiles(inputDir, "*.docx"))
        {
            Document doc = new Document(docFilePath);
            NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;

            foreach (Shape shape in shapeNodes.OfType<Shape>())
            {
                if (!shape.HasImage)
                    continue;

                // Determine image type and appropriate file extension.
                ImageType imgType = shape.ImageData.ImageType;
                string extension = FileFormatUtil.ImageTypeToExtension(imgType); // includes leading dot.

                // Build deterministic image file name.
                string imageFileName = $"{Path.GetFileNameWithoutExtension(docFilePath)}_Image{imageIndex}{extension}";
                string imageFullPath = Path.Combine(imageOutputDir, imageFileName);

                // Save the image to disk.
                shape.ImageData.Save(imageFullPath);
                anyImageExtracted = true;

                // Retrieve image size information.
                ImageSize imgSize = shape.ImageData.ImageSize;

                // Prepare CSV line.
                string csvLine = string.Join(",",
                    Path.GetFileName(docFilePath),
                    imageFileName,
                    imgType,
                    imgSize.WidthPoints,
                    imgSize.HeightPoints,
                    imgSize.WidthPixels,
                    imgSize.HeightPixels,
                    imgSize.HorizontalResolution,
                    imgSize.VerticalResolution);

                csvLines.Add(csvLine);
                imageIndex++;
            }
        }

        // Validate that at least one image was extracted.
        if (!anyImageExtracted)
            throw new InvalidOperationException("No images were extracted from the processed documents.");

        // ---------------------------------------------------------------
        // 4. Write CSV report.
        // ---------------------------------------------------------------
        File.WriteAllLines(csvReportPath, csvLines);

        Console.WriteLine($"Processing complete. Extracted images are in: {imageOutputDir}");
        Console.WriteLine($"CSV report generated at: {csvReportPath}");
    }

    // Helper method to create a deterministic PNG image using Aspose.Drawing.
    private static void CreateSamplePng(string filePath, int width, int height, Aspose.Drawing.Color fillColor)
    {
        // Create bitmap.
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            // Obtain graphics object.
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                // Fill background with the specified color.
                graphics.Clear(fillColor);
            }

            // Save bitmap as PNG.
            bitmap.Save(filePath, ImageFormat.Png);
        }
    }
}
