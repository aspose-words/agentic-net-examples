using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Base directories
        string baseDir = AppDomain.CurrentDomain.BaseDirectory;
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string outputDir = Path.Combine(baseDir, "ExtractedImages");

        // Ensure clean environment
        if (Directory.Exists(inputDir)) Directory.Delete(inputDir, true);
        if (Directory.Exists(outputDir)) Directory.Delete(outputDir, true);
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create sample images of different formats
        CreateSampleImage("sample1.png", 200, 100, Color.LightBlue, ImageFormat.Png);
        CreateSampleImage("sample2.jpg", 150, 150, Color.LightCoral, ImageFormat.Jpeg);
        CreateSampleImage("sample3.gif", 120, 180, Color.LightGreen, ImageFormat.Gif);

        // Create sample DOCX files containing the images
        CreateSampleDocument(Path.Combine(inputDir, "DocumentA.docx"));
        CreateSampleDocument(Path.Combine(inputDir, "DocumentB.docx"));

        // Process each DOCX file: extract images and organize by format
        int totalExtracted = 0;
        foreach (string docPath in Directory.GetFiles(inputDir, "*.docx"))
        {
            Document doc = new Document(docPath);
            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;

            foreach (Shape shape in shapes)
            {
                if (!shape.HasImage) continue;

                ImageData imgData = shape.ImageData;
                string formatFolder = GetFormatFolderName(imgData.ImageType);
                string formatDir = Path.Combine(outputDir, formatFolder);
                Directory.CreateDirectory(formatDir);

                string outputFileName = $"{Path.GetFileNameWithoutExtension(docPath)}_img{imageIndex}.{formatFolder}";
                string outputPath = Path.Combine(formatDir, outputFileName);
                imgData.Save(outputPath);
                imageIndex++;
                totalExtracted++;
            }
        }

        // Validation: ensure at least one image was extracted
        if (totalExtracted == 0)
        {
            throw new InvalidOperationException("No images were extracted from the documents.");
        }

        Console.WriteLine($"Extraction complete. Total images extracted: {totalExtracted}");
    }

    // Helper to create a deterministic sample image file
    private static void CreateSampleImage(string fileName, int width, int height, Color backColor, ImageFormat format)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(backColor);
            }
            bitmap.Save(fileName, format);
        }
    }

    // Helper to create a sample document with the three images inserted
    private static void CreateSampleDocument(string docPath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert PNG
        builder.InsertImage("sample1.png");
        builder.Writeln();

        // Insert JPEG
        builder.InsertImage("sample2.jpg");
        builder.Writeln();

        // Insert GIF
        builder.InsertImage("sample3.gif");
        builder.Writeln();

        doc.Save(docPath);
    }

    // Map Aspose.Words.ImageType to a folder name / file extension
    private static string GetFormatFolderName(ImageType imageType)
    {
        switch (imageType)
        {
            case ImageType.Jpeg:
                return "jpeg";
            case ImageType.Png:
                return "png";
            case ImageType.Gif:
                return "gif";
            case ImageType.Bmp:
                return "bmp";
            case ImageType.Emf:
                return "emf";
            case ImageType.Wmf:
                return "wmf";
            default:
                return "other";
        }
    }
}
