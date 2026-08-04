using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Newtonsoft.Json;

public class Program
{
    public static void Main(string[] args)
    {
        // Define folders and files
        string baseDir = Directory.GetCurrentDirectory();
        string inputFolder = Path.Combine(baseDir, "InputDocs");
        string imageOutputFolder = Path.Combine(baseDir, "ExtractedImages");
        string reportPath = Path.Combine(baseDir, "ImageReport.csv");

        // Prepare folders
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(imageOutputFolder);

        // Create sample images
        string sampleImage1 = Path.Combine(baseDir, "sample1.png");
        string sampleImage2 = Path.Combine(baseDir, "sample2.png");
        CreateSamplePng(sampleImage1, 200, 150, Aspose.Drawing.Color.LightBlue);
        CreateSamplePng(sampleImage2, 120, 180, Aspose.Drawing.Color.LightCoral);

        // Create sample DOCX files with images
        for (int i = 1; i <= 3; i++)
        {
            string docPath = Path.Combine(inputFolder, $"Doc{i}.docx");
            CreateSampleDocumentWithImages(docPath, new[] { sampleImage1, sampleImage2 });
        }

        // Process documents: extract images and collect metadata
        var records = new List<ImageRecord>();
        foreach (string docFile in Directory.GetFiles(inputFolder, "*.docx"))
        {
            Document doc = new Document(docFile);
            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 1;
            foreach (Shape shape in shapes)
            {
                if (shape.HasImage)
                {
                    // Determine image file extension based on image type
                    string ext = GetExtensionFromImageType(shape.ImageData.ImageType);
                    string imageFileName = $"{Path.GetFileNameWithoutExtension(docFile)}_Image{imageIndex}{ext}";
                    string imagePath = Path.Combine(imageOutputFolder, imageFileName);

                    // Save image
                    shape.ImageData.Save(imagePath);

                    // Gather metadata
                    FileInfo fi = new FileInfo(imagePath);
                    var record = new ImageRecord
                    {
                        DocumentName = Path.GetFileName(docFile),
                        ImageFileName = imageFileName,
                        ImageFormat = shape.ImageData.ImageType.ToString(),
                        ImageSizeBytes = fi.Length,
                        WidthPoints = shape.Width,
                        HeightPoints = shape.Height
                    };
                    records.Add(record);
                    imageIndex++;
                }
            }
        }

        // Validate that images were extracted
        if (records.Count == 0)
        {
            throw new InvalidOperationException("No images were extracted from the documents.");
        }

        // Write CSV report
        using (var writer = new StreamWriter(reportPath, false, System.Text.Encoding.UTF8))
        {
            writer.WriteLine("DocumentName,ImageFileName,ImageFormat,ImageSizeBytes,WidthPoints,HeightPoints");
            foreach (var rec in records)
            {
                writer.WriteLine(string.Join(",",
                    EscapeCsv(rec.DocumentName),
                    EscapeCsv(rec.ImageFileName),
                    EscapeCsv(rec.ImageFormat),
                    rec.ImageSizeBytes.ToString(CultureInfo.InvariantCulture),
                    rec.WidthPoints.ToString(CultureInfo.InvariantCulture),
                    rec.HeightPoints.ToString(CultureInfo.InvariantCulture)));
            }
        }

        // Simple validation that report was created
        if (!File.Exists(reportPath) || new FileInfo(reportPath).Length == 0)
        {
            throw new InvalidOperationException("CSV report was not created successfully.");
        }
    }

    private static void CreateSamplePng(string path, int width, int height, Aspose.Drawing.Color backColor)
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

    private static void CreateSampleDocumentWithImages(string docPath, string[] imagePaths)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln($"Sample document: {Path.GetFileName(docPath)}");
        foreach (string imgPath in imagePaths)
        {
            builder.InsertImage(imgPath);
            builder.Writeln(); // add a line break after each image
        }
        doc.Save(docPath);
    }

    private static string GetExtensionFromImageType(ImageType type)
    {
        switch (type)
        {
            case ImageType.Jpeg: return ".jpg";
            case ImageType.Png: return ".png";
            case ImageType.Gif: return ".gif";
            case ImageType.Bmp: return ".bmp";
            case ImageType.Emf: return ".emf";
            case ImageType.Wmf: return ".wmf";
            default: return ".img";
        }
    }

    private static string EscapeCsv(string field)
    {
        if (field.Contains(",") || field.Contains("\"") || field.Contains("\n"))
        {
            return $"\"{field.Replace("\"", "\"\"")}\"";
        }
        return field;
    }

    private class ImageRecord
    {
        public string DocumentName { get; set; }
        public string ImageFileName { get; set; }
        public string ImageFormat { get; set; }
        public long ImageSizeBytes { get; set; }
        public double WidthPoints { get; set; }
        public double HeightPoints { get; set; }
    }
}
