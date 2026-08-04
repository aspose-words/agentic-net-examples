using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Words.Loading;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Base working directory.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        string inputDocsDir = Path.Combine(baseDir, "InputDocs");
        string extractedImagesDir = Path.Combine(baseDir, "ExtractedImages");
        string thumbnailsDir = Path.Combine(baseDir, "Thumbnails");
        string outputDir = Path.Combine(baseDir, "Output");

        // Ensure all directories exist.
        Directory.CreateDirectory(inputDocsDir);
        Directory.CreateDirectory(extractedImagesDir);
        Directory.CreateDirectory(thumbnailsDir);
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create deterministic sample images (input.png, input2.png).
        // -----------------------------------------------------------------
        string sampleImagePath1 = Path.Combine(baseDir, "sample1.png");
        string sampleImagePath2 = Path.Combine(baseDir, "sample2.png");

        CreateSampleImage(sampleImagePath1, 200, 200, Aspose.Drawing.Color.LightBlue, "Img1");
        CreateSampleImage(sampleImagePath2, 200, 200, Aspose.Drawing.Color.LightCoral, "Img2");

        // -----------------------------------------------------------------
        // 2. Create sample DOCX files that contain the images.
        // -----------------------------------------------------------------
        for (int i = 1; i <= 2; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln($"Document {i}");
            builder.InsertParagraph();

            // Insert both sample images into each document.
            builder.InsertImage(sampleImagePath1);
            builder.InsertParagraph();
            builder.InsertImage(sampleImagePath2);
            builder.InsertParagraph();

            string docPath = Path.Combine(inputDocsDir, $"SampleDoc{i}.docx");
            doc.Save(docPath);
        }

        // -----------------------------------------------------------------
        // 3. Batch process: extract images from each DOCX and create thumbnails.
        // -----------------------------------------------------------------
        var docFiles = Directory.GetFiles(inputDocsDir, "*.docx");
        foreach (var docFile in docFiles)
        {
            Document doc = new Document(docFile);
            NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;

            foreach (Shape shape in shapeNodes.OfType<Shape>())
            {
                if (!shape.HasImage) continue;

                // Determine file extension based on image type.
                string ext = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string imageFileName = $"{Path.GetFileNameWithoutExtension(docFile)}_img{imageIndex}{ext}";
                string imagePath = Path.Combine(extractedImagesDir, imageFileName);

                // Save the extracted image.
                shape.ImageData.Save(imagePath);
                if (!File.Exists(imagePath))
                    throw new InvalidOperationException($"Failed to save extracted image: {imagePath}");

                // Create a thumbnail (max width 100px, preserve aspect ratio).
                using (Bitmap originalBmp = new Bitmap(imagePath))
                {
                    int thumbWidth = 100;
                    int thumbHeight = (int)(originalBmp.Height * (thumbWidth / (double)originalBmp.Width));

                    using (Bitmap thumbBmp = new Bitmap(thumbWidth, thumbHeight))
                    {
                        using (Graphics g = Graphics.FromImage(thumbBmp))
                        {
                            g.Clear(Aspose.Drawing.Color.White);
                            g.DrawImage(originalBmp, 0, 0, thumbWidth, thumbHeight);
                        }

                        string thumbFileName = $"{Path.GetFileNameWithoutExtension(imageFileName)}_thumb{ext}";
                        string thumbPath = Path.Combine(thumbnailsDir, thumbFileName);
                        thumbBmp.Save(thumbPath);
                        if (!File.Exists(thumbPath))
                            throw new InvalidOperationException($"Failed to save thumbnail: {thumbPath}");
                    }
                }

                imageIndex++;
            }
        }

        // -----------------------------------------------------------------
        // 4. Build a PDF catalog that contains all thumbnails.
        // -----------------------------------------------------------------
        Document catalog = new Document();
        DocumentBuilder catalogBuilder = new DocumentBuilder(catalog);
        catalogBuilder.Writeln("Image Catalog");
        catalogBuilder.InsertParagraph();

        var thumbFiles = Directory.GetFiles(thumbnailsDir)
                                  .OrderBy(f => f)
                                  .ToArray();

        if (thumbFiles.Length == 0)
            throw new InvalidOperationException("No thumbnails were generated.");

        foreach (var thumbFile in thumbFiles)
        {
            // Add a caption with the thumbnail file name (without extension).
            string caption = Path.GetFileNameWithoutExtension(thumbFile);
            catalogBuilder.Writeln(caption);
            catalogBuilder.InsertImage(thumbFile);
            catalogBuilder.InsertParagraph();
        }

        // Save the catalog as PDF with default options.
        string catalogPath = Path.Combine(outputDir, "ImageCatalog.pdf");
        catalog.Save(catalogPath, SaveFormat.Pdf);
        if (!File.Exists(catalogPath))
            throw new InvalidOperationException($"Failed to save PDF catalog: {catalogPath}");

        // -----------------------------------------------------------------
        // 5. Clean up resources (dispose bitmaps already handled via using).
        // -----------------------------------------------------------------
        // All work completed successfully.
    }

    // Helper method to create a deterministic sample image.
    private static void CreateSampleImage(string path, int width, int height, Aspose.Drawing.Color backColor, string text)
    {
        using (Bitmap bmp = new Bitmap(width, height))
        {
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.Clear(backColor);
                // Simple text drawing using Aspose.Drawing.Font.
                using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 24))
                {
                    g.DrawString(text, font, new SolidBrush(Aspose.Drawing.Color.Black), new PointF(10, height / 2 - 12));
                }
            }
            bmp.Save(path);
        }
    }
}
