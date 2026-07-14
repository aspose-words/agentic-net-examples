using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Prepare sample image
        const string sampleImagePath = "sample.png";
        CreateSampleImage(sampleImagePath, 200, 150, Aspose.Drawing.Color.LightBlue, "Sample");

        // Create a DOCX with a few images
        const string docPath = "sample.docx";
        CreateDocumentWithImages(docPath, sampleImagePath, 3);

        // Load the document and extract images
        Document doc = new Document(docPath);
        List<ImageInfo> extractedImages = ExtractImages(doc, "ExtractedImages");

        // Validate extraction
        if (extractedImages.Count == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        // Generate CSV (Excel-readable) metadata file
        const string csvPath = "ImageMetadata.csv";
        GenerateCsvMetadata(extractedImages, csvPath);

        // Validate CSV creation
        if (!File.Exists(csvPath) || new FileInfo(csvPath).Length == 0)
            throw new InvalidOperationException("Failed to create the image metadata CSV file.");

        // Program completed
        Console.WriteLine("Image extraction and metadata export completed successfully.");
    }

    private static void CreateSampleImage(string path, int width, int height, Aspose.Drawing.Color backColor, string text)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(backColor);
                // Simple rectangle to have deterministic content
                using (Pen pen = new Pen(Aspose.Drawing.Color.Black, 2))
                {
                    g.DrawRectangle(pen, 10, 10, width - 20, height - 20);
                }
            }
            bitmap.Save(path, ImageFormat.Png);
        }
    }

    private static void CreateDocumentWithImages(string docPath, string imagePath, int repeatCount)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        for (int i = 0; i < repeatCount; i++)
        {
            builder.Writeln($"Image #{i + 1}");
            builder.InsertImage(imagePath);
            builder.Writeln(); // Add spacing
        }

        doc.Save(docPath);
    }

    private static List<ImageInfo> ExtractImages(Document doc, string outputFolder)
    {
        if (!Directory.Exists(outputFolder))
            Directory.CreateDirectory(outputFolder);

        List<ImageInfo> images = new List<ImageInfo>();
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

        int index = 1;
        foreach (Shape shape in shapes)
        {
            if (!shape.HasImage)
                continue;

            string ext = shape.ImageData.ImageType.ToString().ToLower(); // e.g., png, jpeg
            string fileName = $"image_{index}.{ext}";
            string fullPath = Path.Combine(outputFolder, fileName);

            // Save image
            shape.ImageData.Save(fullPath);

            // Get metadata
            long fileSize = new FileInfo(fullPath).Length;
            int width, height;
            using (Bitmap bmp = new Bitmap(fullPath))
            {
                width = bmp.Width;
                height = bmp.Height;
            }

            images.Add(new ImageInfo
            {
                FileName = fileName,
                FilePath = fullPath,
                FileSizeBytes = fileSize,
                Width = width,
                Height = height,
                ImageFormat = ext.ToUpperInvariant()
            });

            index++;
        }

        return images;
    }

    private static void GenerateCsvMetadata(List<ImageInfo> images, string csvPath)
    {
        var sb = new StringBuilder();
        sb.AppendLine("FileName,FilePath,FileSizeBytes,Width,Height,ImageFormat");

        foreach (var img in images)
        {
            sb.AppendLine($"{Escape(img.FileName)},{Escape(img.FilePath)},{img.FileSizeBytes},{img.Width},{img.Height},{Escape(img.ImageFormat)}");
        }

        File.WriteAllText(csvPath, sb.ToString(), Encoding.UTF8);
    }

    private static string Escape(string value)
    {
        if (value.Contains(",") || value.Contains("\""))
        {
            value = value.Replace("\"", "\"\"");
            return $"\"{value}\"";
        }
        return value;
    }

    private class ImageInfo
    {
        public string FileName { get; set; }
        public string FilePath { get; set; }
        public long FileSizeBytes { get; set; }
        public int Width { get; set; }
        public int Height { get; set; }
        public string ImageFormat { get; set; }
    }
}
