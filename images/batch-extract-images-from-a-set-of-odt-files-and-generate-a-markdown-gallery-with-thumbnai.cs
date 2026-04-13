using System;
using System.IO;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Base folders
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string imagesDir = Path.Combine(baseDir, "OutputImages");
        string thumbsDir = Path.Combine(baseDir, "Thumbnails");
        string markdownPath = Path.Combine(baseDir, "gallery.md");

        // Ensure folders exist
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(imagesDir);
        Directory.CreateDirectory(thumbsDir);

        // 1. Create a deterministic sample image (sample.png)
        string sampleImagePath = Path.Combine(baseDir, "sample.png");
        CreateSampleImage(sampleImagePath, 200, 200);

        // 2. Create a few sample ODT documents that contain the sample image
        CreateSampleOdtDocuments(inputDir, sampleImagePath, 2);

        // 3. Extract images from each ODT, create thumbnails, and build markdown
        var markdownBuilder = new StringBuilder();
        int totalExtracted = 0;

        foreach (string odtFile in Directory.GetFiles(inputDir, "*.odt"))
        {
            Document doc = new Document(odtFile);
            var shapes = doc.GetChildNodes(NodeType.Shape, true)
                            .Cast<Shape>()
                            .Where(s => s.HasImage)
                            .ToArray();

            if (shapes.Length == 0)
                continue; // No images in this document

            markdownBuilder.AppendLine($"## {Path.GetFileName(odtFile)}");
            markdownBuilder.AppendLine();

            int imageIndex = 0;
            foreach (Shape shape in shapes)
            {
                // Determine file extension based on image type
                string extension = Aspose.Words.FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string baseName = $"{Path.GetFileNameWithoutExtension(odtFile)}_img{imageIndex}{extension}";
                string imagePath = Path.Combine(imagesDir, baseName);

                // Save the full‑size image
                shape.ImageData.Save(imagePath);
                totalExtracted++;

                // Create thumbnail
                string thumbPath = Path.Combine(thumbsDir, baseName);
                CreateThumbnail(imagePath, thumbPath, 150);

                // Add markdown entry (thumbnail links to full image)
                string thumbRelative = Path.Combine("Thumbnails", baseName).Replace("\\", "/");
                string imageRelative = Path.Combine("OutputImages", baseName).Replace("\\", "/");
                markdownBuilder.AppendLine($"[![{baseName}]({thumbRelative})]({imageRelative})");
                markdownBuilder.AppendLine();

                imageIndex++;
            }

            markdownBuilder.AppendLine();
        }

        // Validation: at least one image must have been extracted
        if (totalExtracted == 0)
            throw new InvalidOperationException("No images were extracted from the ODT files.");

        // Write markdown gallery
        File.WriteAllText(markdownPath, markdownBuilder.ToString());
    }

    // Creates a deterministic sample PNG image using Aspose.Drawing.
    private static void CreateSampleImage(string filePath, int width, int height)
    {
        // Ensure any existing file is overwritten
        if (File.Exists(filePath))
            File.Delete(filePath);

        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height);
        Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap);
        graphics.Clear(Aspose.Drawing.Color.LightGray);
        // Simple deterministic drawing (a filled rectangle)
        graphics.FillRectangle(new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.DarkBlue), 20, 20, width - 40, height - 40);
        graphics.Dispose();
        bitmap.Save(filePath);
        bitmap.Dispose();
    }

    // Generates a number of ODT documents each containing the sample image.
    private static void CreateSampleOdtDocuments(string folder, string imagePath, int count)
    {
        for (int i = 1; i <= count; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"Sample ODT Document {i}");
            // Insert the sample image twice to have multiple images per document
            builder.InsertImage(imagePath);
            builder.InsertParagraph();
            builder.InsertImage(imagePath);
            string odtPath = Path.Combine(folder, $"SampleDocument{i}.odt");
            doc.Save(odtPath, SaveFormat.Odt);
        }
    }

    // Generates a thumbnail for a given image file using Aspose.Drawing.
    private static void CreateThumbnail(string sourcePath, string thumbPath, int maxWidth)
    {
        using (Aspose.Drawing.Image sourceImage = Aspose.Drawing.Image.FromFile(sourcePath))
        {
            int thumbWidth = maxWidth;
            int thumbHeight = (int)(sourceImage.Height * (thumbWidth / (double)sourceImage.Width));

            Aspose.Drawing.Bitmap thumbBitmap = new Aspose.Drawing.Bitmap(thumbWidth, thumbHeight);
            Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(thumbBitmap);
            g.Clear(Aspose.Drawing.Color.White);
            g.DrawImage(sourceImage, 0, 0, thumbWidth, thumbHeight);
            g.Dispose();

            using (MemoryStream ms = new MemoryStream())
            {
                thumbBitmap.Save(ms, ImageFormat.Png);
                ms.Position = 0;
                using (FileStream fileStream = new FileStream(thumbPath, FileMode.Create, FileAccess.Write))
                {
                    ms.CopyTo(fileStream);
                }
            }

            thumbBitmap.Dispose();
        }
    }
}
