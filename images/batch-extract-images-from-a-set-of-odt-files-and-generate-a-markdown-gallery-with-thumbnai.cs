using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define deterministic folders
        string baseDir = Directory.GetCurrentDirectory();
        string inputFolder = Path.Combine(baseDir, "InputDocs");
        string imageFolder = Path.Combine(baseDir, "ExtractedImages");
        string thumbFolder = Path.Combine(baseDir, "Thumbnails");
        string outputFolder = Path.Combine(baseDir, "Output");

        // Ensure folders exist
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(imageFolder);
        Directory.CreateDirectory(thumbFolder);
        Directory.CreateDirectory(outputFolder);

        // Step 1: Create a sample image that will be inserted into ODT files
        string sampleImagePath = Path.Combine(baseDir, "sample.png");
        CreateSampleImage(sampleImagePath, 200, 200, Aspose.Drawing.Color.LightBlue, "Sample");

        // Step 2: Generate a few ODT documents containing the sample image
        GenerateSampleOdtDocuments(inputFolder, sampleImagePath, 3);

        // Step 3: Process each ODT, extract images, create thumbnails, and collect markdown entries
        List<string> markdownLines = new List<string>();
        int extractedCount = 0;

        foreach (string odtPath in Directory.GetFiles(inputFolder, "*.odt"))
        {
            Document doc = new Document(odtPath);
            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

            foreach (Shape shape in shapes)
            {
                if (!shape.HasImage) continue;

                // Build deterministic file names
                string imageFileName = $"image_{extractedCount + 1}.png";
                string thumbFileName = $"thumb_{extractedCount + 1}.png";

                string imagePath = Path.Combine(imageFolder, imageFileName);
                string thumbPath = Path.Combine(thumbFolder, thumbFileName);

                // Save original image
                shape.ImageData.Save(imagePath);

                // Create thumbnail (100x100)
                CreateThumbnail(imagePath, thumbPath, 100, 100);

                // Add markdown entry
                string relativeImagePath = Path.GetRelativePath(outputFolder, imagePath).Replace('\\', '/');
                string relativeThumbPath = Path.GetRelativePath(outputFolder, thumbPath).Replace('\\', '/');
                markdownLines.Add($"[![Thumbnail]({relativeThumbPath})]({relativeImagePath})");

                extractedCount++;
            }
        }

        // Validation: ensure at least one image was extracted
        if (extractedCount == 0)
            throw new InvalidOperationException("No images were extracted from the ODT files.");

        // Step 4: Write markdown gallery file
        string markdownPath = Path.Combine(outputFolder, "gallery.md");
        File.WriteAllLines(markdownPath, markdownLines);

        Console.WriteLine($"Extraction complete. {extractedCount} images processed.");
        Console.WriteLine($"Markdown gallery generated at: {markdownPath}");
    }

    // Creates a deterministic sample PNG image using Aspose.Drawing
    private static void CreateSampleImage(string path, int width, int height, Aspose.Drawing.Color backColor, string text)
    {
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height))
        {
            using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                g.Clear(backColor);

                // Fill rectangle
                using (Aspose.Drawing.SolidBrush brush = new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.DarkBlue))
                {
                    g.FillRectangle(brush, 10, 10, width - 20, height - 20);
                }

                // Draw text
                using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 20))
                {
                    using (Aspose.Drawing.SolidBrush textBrush = new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.White))
                    {
                        g.DrawString(text, font, textBrush, new Aspose.Drawing.PointF(20, height / 2 - 15));
                    }
                }
            }

            // Save the bitmap as PNG
            bitmap.Save(path, Aspose.Drawing.Imaging.ImageFormat.Png);
        }
    }

    // Generates a given number of ODT documents each containing the sample image
    private static void GenerateSampleOdtDocuments(string folder, string imagePath, int count)
    {
        for (int i = 1; i <= count; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln($"Document {i}");
            builder.InsertParagraph();
            builder.InsertImage(imagePath);
            builder.Writeln($"End of Document {i}");

            string odtPath = Path.Combine(folder, $"Sample_{i}.odt");
            doc.Save(odtPath, SaveFormat.Odt);
        }
    }

    // Creates a thumbnail from an existing image file
    private static void CreateThumbnail(string sourcePath, string thumbPath, int thumbWidth, int thumbHeight)
    {
        using (Aspose.Drawing.Bitmap sourceBitmap = new Aspose.Drawing.Bitmap(sourcePath))
        {
            using (Aspose.Drawing.Bitmap thumbBitmap = new Aspose.Drawing.Bitmap(thumbWidth, thumbHeight))
            {
                using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(thumbBitmap))
                {
                    g.Clear(Aspose.Drawing.Color.White);
                    g.DrawImage(sourceBitmap, new Aspose.Drawing.Rectangle(0, 0, thumbWidth, thumbHeight));
                }

                thumbBitmap.Save(thumbPath, Aspose.Drawing.Imaging.ImageFormat.Png);
            }
        }
    }
}
