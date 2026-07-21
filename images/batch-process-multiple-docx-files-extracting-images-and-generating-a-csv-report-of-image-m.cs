using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;

public class BatchImageExtractor
{
    public static void Main()
    {
        // Root folder for the demo.
        string rootFolder = Path.Combine(Directory.GetCurrentDirectory(), "BatchImagesDemo");
        string inputFolder = Path.Combine(rootFolder, "InputDocs");
        string imageFolder = Path.Combine(rootFolder, "ExtractedImages");
        string outputFolder = Path.Combine(rootFolder, "Output");
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(imageFolder);
        Directory.CreateDirectory(outputFolder);

        // Create deterministic sample images.
        string sampleImage1 = Path.Combine(rootFolder, "sample1.png");
        string sampleImage2 = Path.Combine(rootFolder, "sample2.png");
        CreateSampleImage(sampleImage1, 120, 80, Aspose.Drawing.Color.LightBlue);
        CreateSampleImage(sampleImage2, 80, 120, Aspose.Drawing.Color.LightCoral);

        // Create a few sample DOCX files that contain the images.
        for (int docIndex = 1; docIndex <= 3; docIndex++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"Document {docIndex}");
            // Insert both sample images.
            builder.InsertImage(sampleImage1);
            builder.InsertParagraph();
            builder.InsertImage(sampleImage2);
            string docPath = Path.Combine(inputFolder, $"Document{docIndex}.docx");
            doc.Save(docPath);
        }

        // Prepare CSV report.
        var csvLines = new List<string>();
        csvLines.Add("DocumentName,ImageFileName,ImageType,WidthPoints,HeightPoints,WidthPixels,HeightPixels,HorizontalResolution,VerticalResolution");

        int totalImageCount = 0;

        // Process each DOCX file in the input folder.
        foreach (string docPath in Directory.GetFiles(inputFolder, "*.docx"))
        {
            Document doc = new Document(docPath);
            NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;

            foreach (Shape shape in shapeNodes.OfType<Shape>())
            {
                if (!shape.HasImage)
                    continue;

                // Determine file extension based on image type.
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string imageFileName = $"{Path.GetFileNameWithoutExtension(docPath)}_Image{imageIndex}{extension}";
                string imageFullPath = Path.Combine(imageFolder, imageFileName);

                // Save the image.
                shape.ImageData.Save(imageFullPath);

                // Gather metadata.
                ImageSize size = shape.ImageData.ImageSize;
                string line = string.Join(",",
                    Path.GetFileName(docPath),
                    imageFileName,
                    shape.ImageData.ImageType,
                    shape.Width,
                    shape.Height,
                    size.WidthPixels,
                    size.HeightPixels,
                    size.HorizontalResolution,
                    size.VerticalResolution);

                csvLines.Add(line);
                imageIndex++;
                totalImageCount++;
            }
        }

        // Validate that at least one image was extracted.
        if (totalImageCount == 0)
            throw new InvalidOperationException("No images were extracted from the documents.");

        // Write CSV report.
        string csvPath = Path.Combine(outputFolder, "ImageMetadataReport.csv");
        File.WriteAllLines(csvPath, csvLines, Encoding.UTF8);
    }

    // Helper to create a deterministic bitmap image.
    private static void CreateSampleImage(string filePath, int width, int height, Aspose.Drawing.Color backColor)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(backColor);
            bitmap.Save(filePath);
        }
    }
}
