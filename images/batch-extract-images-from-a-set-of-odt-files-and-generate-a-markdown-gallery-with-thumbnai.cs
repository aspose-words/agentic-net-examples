using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Prepare folders
        string baseDir = Directory.GetCurrentDirectory();
        string inputFolder = Path.Combine(baseDir, "InputDocs");
        string imageFolder = Path.Combine(baseDir, "ExtractedImages");
        string thumbFolder = Path.Combine(baseDir, "Thumbnails");
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(imageFolder);
        Directory.CreateDirectory(thumbFolder);

        // Create sample images
        string sampleImage1 = Path.Combine(baseDir, "sample1.png");
        string sampleImage2 = Path.Combine(baseDir, "sample2.png");
        CreateSampleImage(sampleImage1, Aspose.Drawing.Color.Red);
        CreateSampleImage(sampleImage2, Aspose.Drawing.Color.Blue);

        // Create sample ODT documents containing the images
        CreateSampleDocument(Path.Combine(inputFolder, "doc1.odt"), sampleImage1);
        CreateSampleDocument(Path.Combine(inputFolder, "doc2.odt"), sampleImage2);

        // Process each ODT file
        var markdownLines = new List<string>();
        int totalExtractedImages = 0;

        foreach (string odtPath in Directory.GetFiles(inputFolder, "*.odt"))
        {
            string docName = Path.GetFileNameWithoutExtension(odtPath);
            markdownLines.Add($"## {docName}");
            Document doc = new Document(odtPath);
            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 1;

            foreach (Shape shape in shapes)
            {
                if (!shape.HasImage) continue;

                // Determine image file name and path
                string imageExt = GetImageExtension(shape.ImageData.ImageType);
                string imageFileName = $"{docName}_img{imageIndex}{imageExt}";
                string imagePath = Path.Combine(imageFolder, imageFileName);

                // Save the extracted image
                shape.ImageData.Save(imagePath);
                totalExtractedImages++;

                // Create thumbnail
                string thumbFileName = $"thumb_{imageFileName}";
                string thumbPath = Path.Combine(thumbFolder, thumbFileName);
                CreateThumbnail(imagePath, thumbPath, 100);

                // Add markdown entry (thumbnail linking to full image)
                string relativeThumb = Path.GetRelativePath(baseDir, thumbPath).Replace("\\", "/");
                string relativeImage = Path.GetRelativePath(baseDir, imagePath).Replace("\\", "/");
                markdownLines.Add($"[![{imageFileName}]({relativeThumb})]({relativeImage})");

                imageIndex++;
            }

            if (imageIndex == 1)
            {
                // No images found in this document
                markdownLines.Add("_No images found in this document._");
            }

            markdownLines.Add(string.Empty); // Blank line for readability
        }

        // Validate that at least one image was extracted
        if (totalExtractedImages == 0)
        {
            throw new InvalidOperationException("No images were extracted from the ODT files.");
        }

        // Write markdown gallery
        string markdownPath = Path.Combine(baseDir, "gallery.md");
        File.WriteAllLines(markdownPath, markdownLines);
        Console.WriteLine($"Gallery generated at: {markdownPath}");
    }

    private static void CreateSampleImage(string path, Aspose.Drawing.Color fillColor)
    {
        int width = 200;
        int height = 200;
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height))
        {
            using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                g.Clear(fillColor);
            }
            bitmap.Save(path);
        }
    }

    private static void CreateSampleDocument(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln($"Document containing image: {Path.GetFileName(imagePath)}");
        builder.InsertImage(imagePath);
        doc.Save(docPath, SaveFormat.Odt);
    }

    private static string GetImageExtension(ImageType imageType)
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
            default:
                return ".png";
        }
    }

    private static void CreateThumbnail(string sourcePath, string thumbPath, int maxDimension)
    {
        using (Aspose.Drawing.Bitmap sourceBitmap = new Aspose.Drawing.Bitmap(sourcePath))
        {
            int thumbWidth, thumbHeight;
            if (sourceBitmap.Width > sourceBitmap.Height)
            {
                thumbWidth = maxDimension;
                thumbHeight = (int)(sourceBitmap.Height * (maxDimension / (float)sourceBitmap.Width));
            }
            else
            {
                thumbHeight = maxDimension;
                thumbWidth = (int)(sourceBitmap.Width * (maxDimension / (float)sourceBitmap.Height));
            }

            using (Aspose.Drawing.Bitmap thumbBitmap = new Aspose.Drawing.Bitmap(thumbWidth, thumbHeight))
            {
                using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(thumbBitmap))
                {
                    g.Clear(Aspose.Drawing.Color.White);
                    g.DrawImage(sourceBitmap, 0, 0, thumbWidth, thumbHeight);
                }
                thumbBitmap.Save(thumbPath);
            }
        }
    }
}
