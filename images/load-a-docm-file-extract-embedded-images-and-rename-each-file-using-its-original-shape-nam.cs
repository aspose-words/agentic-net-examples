using System;
using System.Collections.Generic;
using System.IO;
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
        // Prepare directories
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        Directory.CreateDirectory(workDir);
        string outputDir = Path.Combine(workDir, "Output");
        Directory.CreateDirectory(outputDir);

        // Create sample images
        string imgPath1 = Path.Combine(workDir, "sample1.png");
        string imgPath2 = Path.Combine(workDir, "sample2.png");
        CreateSampleImage(imgPath1, 200, 100, Aspose.Drawing.Color.LightBlue, "Img1");
        CreateSampleImage(imgPath2, 150, 150, Aspose.Drawing.Color.LightGreen, "Img2");

        // Create a DOCM document with the images
        string docPath = Path.Combine(workDir, "sample.docm");
        CreateDocumentWithImages(docPath, new[] { imgPath1, imgPath2 });

        // Load the DOCM and extract images, renaming them using the shape name
        ExtractImagesUsingShapeNames(docPath, outputDir);
    }

    private static void CreateSampleImage(string filePath, int width, int height, Aspose.Drawing.Color backColor, string text)
    {
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height))
        {
            using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                g.Clear(backColor);

                using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Black, 2))
                {
                    g.DrawRectangle(pen, 10, 10, width - 20, height - 20);
                }

                using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 12))
                using (Aspose.Drawing.SolidBrush brush = new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.Black))
                {
                    g.DrawString(text, font, brush, new Aspose.Drawing.PointF(15, height / 2 - 10));
                }
            }

            bitmap.Save(filePath, Aspose.Drawing.Imaging.ImageFormat.Png);
        }
    }

    private static void CreateDocumentWithImages(string docPath, string[] imagePaths)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        for (int i = 0; i < imagePaths.Length; i++)
        {
            // Insert image and obtain the shape
            Shape shape = builder.InsertImage(imagePaths[i]);
            shape.Name = $"ImageShape{i + 1}";
            // Add a paragraph break after each image
            builder.Writeln();
        }

        doc.Save(docPath, SaveFormat.Docm);
    }

    private static void ExtractImagesUsingShapeNames(string docPath, string outputDir)
    {
        Document doc = new Document(docPath);
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        List<string> savedFiles = new List<string>();

        foreach (Shape shape in shapes)
        {
            if (shape.HasImage)
            {
                string shapeName = string.IsNullOrEmpty(shape.Name)
                    ? $"UnnamedShape_{Guid.NewGuid()}"
                    : shape.Name;

                string extension = ImageTypeToExtension(shape.ImageData.ImageType);
                string outPath = Path.Combine(outputDir, $"{shapeName}{extension}");
                shape.ImageData.Save(outPath);
                savedFiles.Add(outPath);
            }
        }

        if (savedFiles.Count == 0)
        {
            throw new InvalidOperationException("No images were extracted from the document.");
        }

        // Optional: write a simple JSON manifest of extracted files
        string manifestPath = Path.Combine(outputDir, "manifest.json");
        File.WriteAllText(manifestPath, JsonConvert.SerializeObject(savedFiles, Formatting.Indented));
    }

    private static string ImageTypeToExtension(ImageType imageType)
    {
        switch (imageType)
        {
            case ImageType.Jpeg:
                return ".jpg";
            case ImageType.Png:
                return ".png";
            case ImageType.Gif:
                return ".gif";
            case ImageType.Bmp:
                return ".bmp";
            case ImageType.Emf:
                return ".emf";
            case ImageType.Wmf:
                return ".wmf";
            default:
                return ".img";
        }
    }
}
