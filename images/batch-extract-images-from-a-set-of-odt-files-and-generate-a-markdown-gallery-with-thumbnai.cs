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
        // Define folders
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string imageDir = Path.Combine(baseDir, "ExtractedImages");
        string thumbDir = Path.Combine(baseDir, "Thumbnails");
        string outputDir = Path.Combine(baseDir, "Output");

        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(imageDir);
        Directory.CreateDirectory(thumbDir);
        Directory.CreateDirectory(outputDir);

        // Create sample images to be inserted into ODT files
        string sampleImage1 = Path.Combine(baseDir, "sample1.png");
        string sampleImage2 = Path.Combine(baseDir, "sample2.png");
        CreateSampleImage(sampleImage1, 200, 200, Aspose.Drawing.Color.LightBlue, "Img1");
        CreateSampleImage(sampleImage2, 200, 200, Aspose.Drawing.Color.LightGreen, "Img2");

        // Create sample ODT documents containing the images
        CreateSampleOdt(Path.Combine(inputDir, "doc1.odt"), new[] { sampleImage1, sampleImage2 });
        CreateSampleOdt(Path.Combine(inputDir, "doc2.odt"), new[] { sampleImage2 });

        // Prepare markdown content
        List<string> markdownLines = new List<string>();
        markdownLines.Add("# Image Gallery");
        markdownLines.Add(string.Empty);

        // Process each ODT file
        foreach (string odtPath in Directory.GetFiles(inputDir, "*.odt"))
        {
            Document doc = new Document(odtPath);
            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;

            foreach (Shape shape in shapes)
            {
                if (!shape.HasImage) continue;

                // Determine image file name
                string imageFileName = $"image_{Path.GetFileNameWithoutExtension(odtPath)}_{imageIndex}.png";
                string imagePath = Path.Combine(imageDir, imageFileName);

                // Save the extracted image
                shape.ImageData.Save(imagePath);
                if (!File.Exists(imagePath))
                    throw new InvalidOperationException($"Failed to save extracted image: {imagePath}");

                // Create thumbnail
                string thumbFileName = $"thumb_{Path.GetFileNameWithoutExtension(imageFileName)}.png";
                string thumbPath = Path.Combine(thumbDir, thumbFileName);
                CreateThumbnail(imagePath, thumbPath, 100, 100);
                if (!File.Exists(thumbPath))
                    throw new InvalidOperationException($"Failed to save thumbnail: {thumbPath}");

                // Add entry to markdown
                string relativeThumb = Path.GetRelativePath(outputDir, thumbPath).Replace("\\", "/");
                string relativeImage = Path.GetRelativePath(outputDir, imagePath).Replace("\\", "/");
                markdownLines.Add($"[![]({relativeThumb})]({relativeImage})");
                markdownLines.Add(string.Empty);

                imageIndex++;
            }
        }

        // Validate that at least one image was extracted
        if (markdownLines.Count <= 2)
            throw new InvalidOperationException("No images were extracted from the ODT files.");

        // Save markdown gallery
        string markdownPath = Path.Combine(outputDir, "gallery.md");
        File.WriteAllLines(markdownPath, markdownLines);
        if (!File.Exists(markdownPath))
            throw new InvalidOperationException($"Failed to write markdown file: {markdownPath}");
    }

    // Creates a deterministic sample PNG image
    private static void CreateSampleImage(string path, int width, int height, Aspose.Drawing.Color backColor, string text)
    {
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height))
        {
            using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                g.Clear(backColor);
                // Simple rectangle for visual distinction
                using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Black, 3))
                {
                    g.DrawRectangle(pen, 10, 10, width - 20, height - 20);
                }
                // Draw text
                using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 24))
                {
                    using (Aspose.Drawing.SolidBrush brush = new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.Black))
                    {
                        g.DrawString(text, font, brush, new Aspose.Drawing.PointF(20, height / 2 - 12));
                    }
                }
            }
            bitmap.Save(path, ImageFormat.Png);
        }
    }

    // Creates a simple ODT document and inserts provided images
    private static void CreateSampleOdt(string docPath, string[] imagePaths)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        foreach (string imgPath in imagePaths)
        {
            if (!File.Exists(imgPath))
                throw new FileNotFoundException($"Image file not found: {imgPath}");
            builder.InsertParagraph();
            builder.InsertImage(imgPath);
        }
        doc.Save(docPath, SaveFormat.Odt);
    }

    // Generates a thumbnail from a source image
    private static void CreateThumbnail(string sourcePath, string thumbPath, int thumbWidth, int thumbHeight)
    {
        using (Aspose.Drawing.Bitmap sourceBitmap = new Aspose.Drawing.Bitmap(sourcePath))
        {
            using (Aspose.Drawing.Bitmap thumbBitmap = new Aspose.Drawing.Bitmap(thumbWidth, thumbHeight))
            {
                using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(thumbBitmap))
                {
                    g.Clear(Aspose.Drawing.Color.White);
                    g.DrawImage(sourceBitmap, 0, 0, thumbWidth, thumbHeight);
                }
                thumbBitmap.Save(thumbPath, ImageFormat.Png);
            }
        }
    }
}
