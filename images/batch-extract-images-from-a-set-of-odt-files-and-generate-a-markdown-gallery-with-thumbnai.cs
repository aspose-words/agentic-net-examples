using System;
using System.IO;
using System.Text;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Base directory of the application.
        string baseDir = Directory.GetCurrentDirectory();

        // Input ODT files, extracted images, thumbnails and markdown gallery paths.
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string imagesDir = Path.Combine(baseDir, "ExtractedImages");
        string thumbsDir = Path.Combine(baseDir, "Thumbnails");
        string markdownPath = Path.Combine(baseDir, "gallery.md");

        // Ensure required folders exist.
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(imagesDir);
        Directory.CreateDirectory(thumbsDir);

        // -----------------------------------------------------------------
        // 1. Create sample ODT documents with embedded images.
        // -----------------------------------------------------------------
        for (int docIndex = 1; docIndex <= 2; docIndex++)
        {
            // Create a deterministic bitmap using Aspose.Drawing.
            string sampleImagePath = Path.Combine(baseDir, $"sample{docIndex}.png");
            using (Bitmap bmp = new Bitmap(200, 200))
            using (Graphics g = Graphics.FromImage(bmp))
            {
                // Fill with a deterministic color.
                g.Clear(docIndex == 1 ? Color.Blue : Color.Green);
                // Save the bitmap to a file.
                bmp.Save(sampleImagePath);
            }

            // Build a document and insert the image.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"Sample document {docIndex}");
            builder.InsertImage(sampleImagePath);
            builder.Writeln("End of document.");

            // Save as ODT.
            string odtPath = Path.Combine(inputDir, $"Sample{docIndex}.odt");
            doc.Save(odtPath, SaveFormat.Odt);
        }

        // -----------------------------------------------------------------
        // 2. Extract images from each ODT and create thumbnails.
        // -----------------------------------------------------------------
        var markdownBuilder = new StringBuilder();
        int totalExtracted = 0;

        foreach (string odtFile in Directory.GetFiles(inputDir, "*.odt"))
        {
            Document doc = new Document(odtFile);
            NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;

            foreach (Shape shape in shapeNodes.OfType<Shape>())
            {
                if (!shape.HasImage) continue;

                // Determine file extension based on image type.
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string baseName = $"{Path.GetFileNameWithoutExtension(odtFile)}_img{imageIndex}{extension}";
                string imagePath = Path.Combine(imagesDir, baseName);

                // Save the original image.
                shape.ImageData.Save(imagePath);
                totalExtracted++;

                // -----------------------------------------------------------------
                // Create a thumbnail (max 150x150 while preserving aspect ratio).
                // -----------------------------------------------------------------
                string thumbName = $"thumb_{baseName}";
                string thumbPath = Path.Combine(thumbsDir, thumbName);

                using (MemoryStream ms = new MemoryStream(File.ReadAllBytes(imagePath)))
                {
                    ms.Position = 0;
                    using (Bitmap original = new Bitmap(ms))
                    {
                        const int maxDim = 150;
                        float ratio = Math.Min((float)maxDim / original.Width, (float)maxDim / original.Height);
                        int thumbWidth = Math.Max(1, (int)(original.Width * ratio));
                        int thumbHeight = Math.Max(1, (int)(original.Height * ratio));

                        using (Bitmap thumb = new Bitmap(thumbWidth, thumbHeight))
                        using (Graphics g = Graphics.FromImage(thumb))
                        {
                            g.Clear(Color.White);
                            g.DrawImage(original, 0, 0, thumbWidth, thumbHeight);
                            thumb.Save(thumbPath);
                        }
                    }
                }

                // -----------------------------------------------------------------
                // Append markdown entry (thumbnail linked to full‑size image).
                // -----------------------------------------------------------------
                string relThumb = Path.GetRelativePath(baseDir, thumbPath).Replace('\\', '/');
                string relImage = Path.GetRelativePath(baseDir, imagePath).Replace('\\', '/');
                markdownBuilder.AppendLine($"[![{baseName}]({relThumb})]({relImage})");

                imageIndex++;
            }
        }

        // Validate that at least one image was extracted.
        if (totalExtracted == 0)
            throw new InvalidOperationException("No images were extracted from the ODT files.");

        // Write the markdown gallery to file.
        File.WriteAllText(markdownPath, markdownBuilder.ToString());

        Console.WriteLine($"Extracted {totalExtracted} images.");
        Console.WriteLine($"Markdown gallery created at: {markdownPath}");
    }
}
