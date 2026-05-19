using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Base directories
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        string inputDir = Path.Combine(baseDir, "Input");
        string outputDir = Path.Combine(baseDir, "Output");
        string thumbsDir = Path.Combine(outputDir, "thumbs");
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);
        Directory.CreateDirectory(thumbsDir);

        // Create a deterministic sample image to be used in ODT files
        string sampleImagePath = Path.Combine(inputDir, "sample.png");
        CreateSampleImage(sampleImagePath, 200, 200, Aspose.Drawing.Color.LightBlue);

        // Create a few ODT documents that contain the sample image
        int docCount = 3;
        for (int i = 1; i <= docCount; i++)
        {
            string odtPath = Path.Combine(inputDir, $"Document{i}.odt");
            CreateOdtWithImage(odtPath, sampleImagePath);
        }

        // Prepare markdown content
        var markdownLines = new System.Collections.Generic.List<string>();
        markdownLines.Add("# Image Gallery");
        markdownLines.Add("");

        int totalExtracted = 0;

        // Process each ODT file
        foreach (string odtFile in Directory.GetFiles(inputDir, "*.odt"))
        {
            Document doc = new Document(odtFile);
            NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;

            foreach (Shape shape in shapeNodes.OfType<Shape>())
            {
                if (!shape.HasImage) continue;

                // Determine file extension based on image type
                string ext = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string baseName = $"{Path.GetFileNameWithoutExtension(odtFile)}_img{imageIndex}{ext}";
                string imagePath = Path.Combine(outputDir, baseName);

                // Save the extracted image
                shape.ImageData.Save(imagePath);
                totalExtracted++;

                // Create thumbnail (max dimension 100px)
                string thumbPath = Path.Combine(thumbsDir, $"{Path.GetFileNameWithoutExtension(baseName)}_thumb{ext}");
                CreateThumbnail(imagePath, thumbPath, 100);

                // Add markdown entry
                string relativeImage = Path.GetFileName(baseName);
                string relativeThumb = Path.Combine("thumbs", Path.GetFileName(thumbPath)).Replace('\\', '/');
                markdownLines.Add($"![Thumbnail]({relativeThumb}) [{relativeImage}]({relativeImage})");
                markdownLines.Add("");

                imageIndex++;
            }
        }

        // Validate that images were extracted
        if (totalExtracted == 0)
            throw new InvalidOperationException("No images were extracted from the ODT files.");

        // Write markdown gallery
        string markdownPath = Path.Combine(outputDir, "gallery.md");
        File.WriteAllLines(markdownPath, markdownLines);
    }

    // Creates a simple bitmap with a solid background color
    private static void CreateSampleImage(string filePath, int width, int height, Aspose.Drawing.Color bgColor)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics g = Graphics.FromImage(bitmap))
        {
            g.Clear(bgColor);
            bitmap.Save(filePath);
        }
    }

    // Creates an ODT document containing a single image
    private static void CreateOdtWithImage(string odtPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln($"Document generated for image extraction: {Path.GetFileName(imagePath)}");
        builder.InsertImage(imagePath);
        doc.Save(odtPath, SaveFormat.Odt);
    }

    // Generates a thumbnail preserving aspect ratio, max dimension defined by maxSize
    private static void CreateThumbnail(string sourcePath, string thumbPath, int maxSize)
    {
        using (Bitmap source = new Bitmap(sourcePath))
        {
            int thumbWidth, thumbHeight;
            if (source.Width > source.Height)
            {
                thumbWidth = maxSize;
                thumbHeight = (int)(source.Height * (maxSize / (float)source.Width));
            }
            else
            {
                thumbHeight = maxSize;
                thumbWidth = (int)(source.Width * (maxSize / (float)source.Height));
            }

            using (Bitmap thumb = new Bitmap(thumbWidth, thumbHeight))
            using (Graphics g = Graphics.FromImage(thumb))
            {
                g.Clear(Aspose.Drawing.Color.White);
                g.DrawImage(source, 0, 0, thumbWidth, thumbHeight);
                thumb.Save(thumbPath);
            }
        }
    }
}
