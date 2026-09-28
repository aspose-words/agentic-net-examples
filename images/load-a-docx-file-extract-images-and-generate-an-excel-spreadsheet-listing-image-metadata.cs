using System;
using System.IO;
using System.Text;
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
        // Create a deterministic sample image file.
        const string sampleImagePath = "sample.png";
        CreateSampleImage(sampleImagePath, 200, 200);

        // Create a DOCX document and insert the sample image.
        const string docPath = "sample.docx";
        CreateDocumentWithImage(docPath, sampleImagePath);

        // Load the DOCX document.
        Document doc = new Document(docPath);

        // Extract images and collect metadata.
        List<ImageInfo> images = ExtractImages(doc, "extracted");

        // Validate that at least one image was extracted.
        if (images.Count == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        // Generate an Excel-compatible CSV file with the metadata.
        const string excelPath = "ImageMetadata.xlsx";
        GenerateExcelCsv(excelPath, images);

        // Validate that the Excel file was created.
        if (!File.Exists(excelPath))
            throw new FileNotFoundException("Failed to create the Excel metadata file.", excelPath);
    }

    private static void CreateSampleImage(string path, int width, int height)
    {
        // Ensure any existing file is removed.
        if (File.Exists(path))
            File.Delete(path);

        // Create bitmap and draw deterministic content.
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height);
        Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap);
        graphics.Clear(Aspose.Drawing.Color.White);
        // Draw a simple rectangle.
        using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Black))
        {
            graphics.DrawRectangle(pen, 10, 10, width - 20, height - 20);
        }

        // Save the bitmap as PNG.
        bitmap.Save(path, Aspose.Drawing.Imaging.ImageFormat.Png);

        // Clean up drawing resources.
        graphics.Dispose();
        bitmap.Dispose();

        // Validate that the image file exists.
        if (!File.Exists(path))
            throw new FileNotFoundException("Failed to create the sample image.", path);
    }

    private static void CreateDocumentWithImage(string docPath, string imagePath)
    {
        // Ensure any existing document is removed.
        if (File.Exists(docPath))
            File.Delete(docPath);

        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath, SaveFormat.Docx);

        // Validate that the document file exists.
        if (!File.Exists(docPath))
            throw new FileNotFoundException("Failed to create the DOCX document.", docPath);
    }

    private static List<ImageInfo> ExtractImages(Document doc, string outputFolder)
    {
        // Ensure output folder exists.
        if (!Directory.Exists(outputFolder))
            Directory.CreateDirectory(outputFolder);

        List<ImageInfo> list = new List<ImageInfo>();
        int index = 1;

        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapes)
        {
            if (!shape.HasImage)
                continue;

            // Determine file extension based on image type.
            string extension = GetExtensionFromImageType(shape.ImageData.ImageType);
            string fileName = $"image-{index}{extension}";
            string fullPath = Path.Combine(outputFolder, fileName);

            // Save the image.
            shape.ImageData.Save(fullPath);

            // Collect metadata.
            ImageInfo info = new ImageInfo
            {
                FileName = fileName,
                Format = shape.ImageData.ImageType.ToString(),
                Width = shape.ImageData.ImageSize.WidthPoints,
                Height = shape.ImageData.ImageSize.HeightPoints,
                SizeBytes = shape.ImageData.ImageBytes.Length
            };
            list.Add(info);
            index++;
        }

        return list;
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

    private static void GenerateExcelCsv(string path, List<ImageInfo> images)
    {
        // Ensure any existing file is removed.
        if (File.Exists(path))
            File.Delete(path);

        StringBuilder sb = new StringBuilder();
        sb.AppendLine("FileName,Format,Width,Height,SizeBytes");
        foreach (var img in images)
        {
            sb.AppendLine($"{img.FileName},{img.Format},{img.Width},{img.Height},{img.SizeBytes}");
        }

        File.WriteAllText(path, sb.ToString(), Encoding.UTF8);
    }

    private class ImageInfo
    {
        public string FileName { get; set; }
        public string Format { get; set; }
        public double Width { get; set; }
        public double Height { get; set; }
        public int SizeBytes { get; set; }
    }
}
