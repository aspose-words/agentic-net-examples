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
        string inputFolder = Path.Combine(baseDir, "InputDocs");
        string imageFolder = Path.Combine(baseDir, "ExtractedImages");
        string thumbFolder = Path.Combine(baseDir, "Thumbnails");
        string markdownPath = Path.Combine(baseDir, "gallery.md");

        // Ensure folders exist
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(imageFolder);
        Directory.CreateDirectory(thumbFolder);

        // Create a deterministic sample image to be used in documents
        string sampleImagePath = Path.Combine(baseDir, "sample.png");
        CreateSampleImage(sampleImagePath, 400, 300);

        // Create sample ODT documents containing the sample image
        CreateSampleDocument(Path.Combine(inputFolder, "doc1.odt"), sampleImagePath, 1);
        CreateSampleDocument(Path.Combine(inputFolder, "doc2.odt"), sampleImagePath, 2);

        // Process each ODT file: extract images, create thumbnails, collect markdown entries
        var markdownLines = new List<string>();
        int totalExtracted = 0;

        foreach (string docPath in Directory.GetFiles(inputFolder, "*.odt"))
        {
            Document doc = new Document(docPath);
            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
            int imgIndex = 0;

            foreach (Shape shape in shapes)
            {
                if (!shape.HasImage) continue;

                string docName = Path.GetFileNameWithoutExtension(docPath);
                string imageFileName = $"{docName}_img{imgIndex}.png";
                string imageFullPath = Path.Combine(imageFolder, imageFileName);

                // Save the extracted image
                shape.ImageData.Save(imageFullPath);
                totalExtracted++;

                // Create thumbnail
                string thumbFileName = $"{docName}_img{imgIndex}_thumb.png";
                string thumbFullPath = Path.Combine(thumbFolder, thumbFileName);
                CreateThumbnail(imageFullPath, thumbFullPath, 150, 150);

                // Add markdown entry
                string relativeImagePath = Path.GetRelativePath(baseDir, imageFullPath).Replace('\\', '/');
                string relativeThumbPath = Path.GetRelativePath(baseDir, thumbFullPath).Replace('\\', '/');
                markdownLines.Add($"[![Thumbnail]({relativeThumbPath})]({relativeImagePath})");

                imgIndex++;
            }
        }

        // Validate that images were extracted
        if (totalExtracted == 0)
            throw new InvalidOperationException("No images were extracted from the ODT files.");

        // Write markdown gallery
        File.WriteAllLines(markdownPath, markdownLines);
    }

    // Creates a simple PNG image with deterministic content
    private static void CreateSampleImage(string path, int width, int height)
    {
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height))
        {
            using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                g.Clear(Aspose.Drawing.Color.White);
                using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Blue, 5))
                {
                    g.DrawRectangle(pen, 10, 10, width - 20, height - 20);
                }
                using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 24))
                {
                    g.DrawString("Sample", font, new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.Black), new Aspose.Drawing.PointF(50, height / 2 - 12));
                }
            }
            bitmap.Save(path, Aspose.Drawing.Imaging.ImageFormat.Png);
        }
    }

    // Creates an ODT document with a given number of inserted images
    private static void CreateSampleDocument(string docPath, string imagePath, int imageCount)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        for (int i = 0; i < imageCount; i++)
        {
            builder.Writeln($"Image {i + 1}:");
            builder.InsertImage(imagePath);
            builder.Writeln(); // Add spacing
        }

        doc.Save(docPath, SaveFormat.Odt);
    }

    // Generates a thumbnail of specified size while preserving aspect ratio
    private static void CreateThumbnail(string sourcePath, string destPath, int maxWidth, int maxHeight)
    {
        using (Aspose.Drawing.Bitmap source = new Aspose.Drawing.Bitmap(sourcePath))
        {
            // Calculate scaling
            double ratioX = (double)maxWidth / source.Width;
            double ratioY = (double)maxHeight / source.Height;
            double ratio = Math.Min(ratioX, ratioY);
            int thumbWidth = (int)(source.Width * ratio);
            int thumbHeight = (int)(source.Height * ratio);

            using (Aspose.Drawing.Bitmap thumb = new Aspose.Drawing.Bitmap(thumbWidth, thumbHeight))
            {
                using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(thumb))
                {
                    g.Clear(Aspose.Drawing.Color.White);
                    g.DrawImage(source, new Aspose.Drawing.Rectangle(0, 0, thumbWidth, thumbHeight));
                }
                thumb.Save(destPath, Aspose.Drawing.Imaging.ImageFormat.Png);
            }
        }
    }
}
