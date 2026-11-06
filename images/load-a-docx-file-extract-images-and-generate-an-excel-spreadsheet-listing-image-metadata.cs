using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare deterministic file names
        string docPath = "sample.docx";
        string image1Path = "input1.png";
        string image2Path = "input2.png";
        string outputFolder = "extracted_images";
        string excelPath = "ImageMetadata.xlsx";

        // Ensure output folder exists
        Directory.CreateDirectory(outputFolder);

        // Create sample images
        CreateSampleImage(image1Path, 200, 150, Aspose.Drawing.Color.LightBlue);
        CreateSampleImage(image2Path, 120, 180, Aspose.Drawing.Color.LightCoral);

        // Create a DOCX and insert the images
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample document with images:");
        builder.InsertImage(image1Path);
        builder.Writeln();
        builder.InsertImage(image2Path);
        doc.Save(docPath);

        // Load the document for extraction
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        var metadataList = new List<ImageMetadata>();

        int imageIndex = 1;
        foreach (Shape shape in shapeNodes)
        {
            if (shape.HasImage)
            {
                // Prepare file name for extracted image
                string extractedImagePath = Path.Combine(outputFolder, $"extracted_{imageIndex}.png");

                // Save the image
                shape.ImageData.Save(extractedImagePath);

                // Get image size in points (1 point = 1/72 inch)
                ImageSize size = shape.ImageData.ImageSize; // WidthPoints / HeightPoints
                // Get file size in bytes
                long fileSize = new FileInfo(extractedImagePath).Length;

                metadataList.Add(new ImageMetadata
                {
                    FileName = Path.GetFileName(extractedImagePath),
                    WidthPoints = size.WidthPoints,
                    HeightPoints = size.HeightPoints,
                    SizeBytes = fileSize
                });

                imageIndex++;
            }
        }

        // Validate that at least one image was extracted
        if (metadataList.Count == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        // Generate CSV content (Excel can open CSV files)
        using (var writer = new StreamWriter(excelPath, false))
        {
            writer.WriteLine("ImageFileName,WidthPoints,HeightPoints,SizeBytes");
            foreach (var meta in metadataList)
            {
                writer.WriteLine($"{meta.FileName},{meta.WidthPoints},{meta.HeightPoints},{meta.SizeBytes}");
            }
        }

        // Validate that the spreadsheet was created
        if (!File.Exists(excelPath))
            throw new InvalidOperationException("Failed to create the Excel spreadsheet.");
    }

    private static void CreateSampleImage(string path, int width, int height, Aspose.Drawing.Color backColor)
    {
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height))
        {
            using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                g.Clear(backColor);
            }
            bitmap.Save(path);
        }
    }

    private class ImageMetadata
    {
        public string FileName { get; set; }
        public double WidthPoints { get; set; }
        public double HeightPoints { get; set; }
        public long SizeBytes { get; set; }
    }
}
