using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Paths for temporary files
        const string imagePath = "input.png";
        const string docPath = "sample.docx";
        const string reportPath = "ImageMetadata.xlsx"; // CSV content with .xlsx extension

        // 1. Create a deterministic sample image
        CreateSampleImage(imagePath);

        // 2. Create a DOCX document and insert the sample image
        CreateDocumentWithImage(docPath, imagePath);

        // 3. Load the DOCX and extract images, collecting metadata
        var metadataList = ExtractImagesAndCollectMetadata(docPath);

        // 4. Generate a spreadsheet (CSV) with the collected metadata
        GenerateCsvReport(reportPath, metadataList);

        // 5. Validate outputs
        ValidateOutputs(metadataList, reportPath);
    }

    private static void CreateSampleImage(string path)
    {
        const int width = 200;
        const int height = 150;

        // Create bitmap using Aspose.Drawing
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height);
        try
        {
            // Obtain graphics object
            Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap);
            try
            {
                // Fill background
                graphics.Clear(Aspose.Drawing.Color.White);
                // Draw a simple rectangle
                using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Black, 3))
                {
                    graphics.DrawRectangle(pen, 10, 10, width - 20, height - 20);
                }
            }
            finally
            {
                graphics.Dispose();
            }

            // Save as PNG
            bitmap.Save(path, Aspose.Drawing.Imaging.ImageFormat.Png);
        }
        finally
        {
            bitmap.Dispose();
        }

        if (!File.Exists(path))
            throw new Exception($"Failed to create sample image at '{path}'.");
    }

    private static void CreateDocumentWithImage(string docPath, string imagePath)
    {
        // Ensure the image exists before insertion
        if (!File.Exists(imagePath))
            throw new FileNotFoundException($"Image file not found: {imagePath}");

        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath);

        if (!File.Exists(docPath))
            throw new Exception($"Failed to save document at '{docPath}'.");
    }

    private static List<ImageMetadata> ExtractImagesAndCollectMetadata(string docPath)
    {
        if (!File.Exists(docPath))
            throw new FileNotFoundException($"Document file not found: {docPath}");

        Document doc = new Document(docPath);
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        List<ImageMetadata> metadataList = new List<ImageMetadata>();
        int imageIndex = 0;

        foreach (Shape shape in shapes)
        {
            if (!shape.HasImage)
                continue;

            string extractedImagePath = $"extracted-{imageIndex}.png";

            // Save the image to a file
            shape.ImageData.Save(extractedImagePath);

            // Load the saved image to obtain dimensions and size
            using (FileStream fs = new FileStream(extractedImagePath, FileMode.Open, FileAccess.Read))
            {
                using (MemoryStream ms = new MemoryStream())
                {
                    fs.CopyTo(ms);
                    ms.Position = 0; // Reset position before reading
                    using (Aspose.Drawing.Bitmap bmp = new Aspose.Drawing.Bitmap(ms))
                    {
                        long sizeBytes = new FileInfo(extractedImagePath).Length;
                        ImageMetadata meta = new ImageMetadata
                        {
                            FileName = extractedImagePath,
                            SizeBytes = sizeBytes,
                            Width = bmp.Width,
                            Height = bmp.Height
                        };
                        metadataList.Add(meta);
                    }
                }
            }

            imageIndex++;
        }

        return metadataList;
    }

    private static void GenerateCsvReport(string reportPath, List<ImageMetadata> metadata)
    {
        using (StreamWriter writer = new StreamWriter(reportPath, false))
        {
            // Header row
            writer.WriteLine("FileName,SizeBytes,Width,Height");

            // Data rows
            foreach (var item in metadata)
            {
                writer.WriteLine($"{item.FileName},{item.SizeBytes},{item.Width},{item.Height}");
            }
        }
    }

    private static void ValidateOutputs(List<ImageMetadata> metadata, string reportPath)
    {
        if (metadata.Count == 0)
            throw new Exception("No images were extracted from the document.");

        if (!File.Exists(reportPath))
            throw new Exception($"Report was not created at '{reportPath}'.");
    }

    private class ImageMetadata
    {
        public string FileName { get; set; }
        public long SizeBytes { get; set; }
        public int Width { get; set; }
        public int Height { get; set; }
    }
}
