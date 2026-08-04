using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class BatchImageExtractor
{
    public static void Main()
    {
        // Base working directory.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "BatchImageExtract");
        string inputDir = Path.Combine(baseDir, "Input");
        string imagesDir = Path.Combine(baseDir, "Images");
        string thumbsDir = Path.Combine(baseDir, "Thumbnails");
        string outputDir = Path.Combine(baseDir, "Output");

        // Ensure clean folders.
        foreach (string dir in new[] { inputDir, imagesDir, thumbsDir, outputDir })
        {
            if (Directory.Exists(dir))
                Directory.Delete(dir, true);
            Directory.CreateDirectory(dir);
        }

        // -------------------------------------------------
        // 1. Create sample images (deterministic local files).
        // -------------------------------------------------
        string sampleImg1 = Path.Combine(baseDir, "sample1.png");
        string sampleImg2 = Path.Combine(baseDir, "sample2.png");
        CreateSampleImage(sampleImg1, 200, 150, Aspose.Drawing.Color.LightBlue);
        CreateSampleImage(sampleImg2, 150, 200, Aspose.Drawing.Color.LightCoral);

        // -------------------------------------------------
        // 2. Create sample ODT documents that contain the images.
        // -------------------------------------------------
        for (int docIndex = 1; docIndex <= 2; docIndex++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln($"Document {docIndex} - contains two images.");
            builder.InsertImage(sampleImg1);
            builder.InsertParagraph();
            builder.InsertImage(sampleImg2);

            string odtPath = Path.Combine(inputDir, $"Sample{docIndex}.odt");
            doc.Save(odtPath, SaveFormat.Odt);
        }

        // -------------------------------------------------
        // 3. Process each ODT file: extract images, create thumbnails, build markdown.
        // -------------------------------------------------
        List<string> markdownLines = new List<string>();
        markdownLines.Add("# Image Gallery");
        markdownLines.Add("");

        string[] odtFiles = Directory.GetFiles(inputDir, "*.odt");
        int totalExtracted = 0;

        foreach (string odtFile in odtFiles)
        {
            Document doc = new Document(odtFile);
            NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

            int imageIndex = 0;
            foreach (Shape shape in shapeNodes.OfType<Shape>())
            {
                if (!shape.HasImage)
                    continue;

                // Determine file extension based on image type.
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string baseName = $"{Path.GetFileNameWithoutExtension(odtFile)}_img{imageIndex}{extension}";
                string imagePath = Path.Combine(imagesDir, baseName);

                // Save the extracted image.
                shape.ImageData.Save(imagePath);
                totalExtracted++;

                // Create thumbnail.
                string thumbName = $"{Path.GetFileNameWithoutExtension(baseName)}_thumb{extension}";
                string thumbPath = Path.Combine(thumbsDir, thumbName);
                CreateThumbnail(imagePath, thumbPath, 150);

                // Add markdown entry.
                string relThumb = Path.Combine("Thumbnails", thumbName).Replace("\\", "/");
                string relImage = Path.Combine("Images", baseName).Replace("\\", "/");
                markdownLines.Add($"![{baseName}]({relThumb})({relImage})");
                markdownLines.Add("");

                imageIndex++;
            }
        }

        // Validation: at least one image must have been extracted.
        if (totalExtracted == 0)
            throw new InvalidOperationException("No images were extracted from the ODT files.");

        // -------------------------------------------------
        // 4. Write markdown gallery file.
        // -------------------------------------------------
        string markdownPath = Path.Combine(outputDir, "gallery.md");
        File.WriteAllLines(markdownPath, markdownLines);
    }

    // Creates a deterministic PNG image using Aspose.Drawing.
    private static void CreateSampleImage(string filePath, int width, int height, Aspose.Drawing.Color backColor)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(backColor);
            bitmap.Save(filePath, ImageFormat.Png);
        }
    }

    // Generates a thumbnail for a given image file, preserving aspect ratio.
    private static void CreateThumbnail(string sourcePath, string thumbPath, int maxWidth)
    {
        using (Bitmap original = new Bitmap(sourcePath))
        {
            int thumbWidth = maxWidth;
            int thumbHeight = (int)(original.Height * (thumbWidth / (double)original.Width));
            if (thumbHeight <= 0) thumbHeight = maxWidth;

            using (Bitmap thumbnail = new Bitmap(thumbWidth, thumbHeight))
            using (Graphics g = Graphics.FromImage(thumbnail))
            {
                g.Clear(Aspose.Drawing.Color.White);
                g.DrawImage(original, 0, 0, thumbWidth, thumbHeight);
                thumbnail.Save(thumbPath, ImageFormat.Png);
            }
        }
    }
}
