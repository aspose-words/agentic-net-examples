using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;
using Aspose.Drawing;               // Aspose.Drawing.Common namespace
using Aspose.Drawing.Imaging;       // For ImageFormat if needed

public class Program
{
    // Entry point
    public static void Main()
    {
        // Root folders for all generated data
        string rootFolder = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        string docsFolder = Path.Combine(rootFolder, "Docs");
        string extractedFolder = Path.Combine(rootFolder, "Extracted");
        string thumbsFolder = Path.Combine(rootFolder, "Thumbs");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "Output");

        // Ensure folders exist
        Directory.CreateDirectory(docsFolder);
        Directory.CreateDirectory(extractedFolder);
        Directory.CreateDirectory(thumbsFolder);
        Directory.CreateDirectory(outputFolder);

        // 1. Create a deterministic sample image that will be inserted into the sample DOCX files
        string sampleImagePath = Path.Combine(rootFolder, "sample.png");
        CreateSampleImage(sampleImagePath, 200, 200, Aspose.Drawing.Color.LightBlue);

        // 2. Create a few sample DOCX files that contain the sample image
        int docCount = 3;
        for (int i = 0; i < docCount; i++)
        {
            string docPath = Path.Combine(docsFolder, $"Document_{i}.docx");
            CreateSampleDocumentWithImage(docPath, sampleImagePath, $"Sample document {i}");
        }

        // 3. Process each DOCX: extract images, create thumbnails
        var docFiles = Directory.GetFiles(docsFolder, "*.docx");
        int totalThumbnails = 0;

        foreach (var docFile in docFiles)
        {
            // Load the document (no special load options needed)
            Document doc = new Document(docFile);

            // Get all shape nodes that contain images
            var shapeNodes = doc.GetChildNodes(NodeType.Shape, true)
                                .Cast<Shape>()
                                .Where(s => s.HasImage)
                                .ToList();

            int imageIndex = 0;
            foreach (var shape in shapeNodes)
            {
                // Build deterministic file names
                string baseName = Path.GetFileNameWithoutExtension(docFile);
                string extractedImagePath = Path.Combine(extractedFolder,
                    $"{baseName}_Image_{imageIndex}{FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType)}");

                // Save the image from the shape
                shape.ImageData.Save(extractedImagePath);

                // Create a thumbnail (max 100x100 while preserving aspect ratio)
                string thumbPath = Path.Combine(thumbsFolder,
                    $"{baseName}_Thumb_{imageIndex}.png");
                CreateThumbnail(extractedImagePath, thumbPath, 100, 100);
                totalThumbnails++;

                imageIndex++;
            }
        }

        // Validate that at least one thumbnail was produced
        if (totalThumbnails == 0)
            throw new InvalidOperationException("No thumbnails were generated. Ensure that source documents contain images.");

        // 4. Build a PDF catalog that contains all thumbnails
        Document catalog = new Document();
        DocumentBuilder builder = new DocumentBuilder(catalog);

        // Optional: add a title
        builder.Writeln("Image Catalog");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln();

        // Insert each thumbnail with a caption
        var thumbFiles = Directory.GetFiles(thumbsFolder, "*.png");
        foreach (var thumbFile in thumbFiles)
        {
            // Insert thumbnail image
            Shape thumbShape = builder.InsertImage(thumbFile);
            thumbShape.WrapType = WrapType.Inline;
            thumbShape.Width = 100;   // Fixed width for consistency
            thumbShape.Height = 100;  // Fixed height for consistency

            // Add a caption below the image
            builder.Writeln(Path.GetFileNameWithoutExtension(thumbFile));
            builder.Writeln(); // Add spacing between entries
        }

        // Save the catalog as PDF
        string catalogPath = Path.Combine(outputFolder, "ImageCatalog.pdf");
        catalog.Save(catalogPath, SaveFormat.Pdf);

        Console.WriteLine($"Catalog PDF created at: {catalogPath}");
    }

    // Creates a simple solid‑color PNG image using Aspose.Drawing
    private static void CreateSampleImage(string filePath, int width, int height, Aspose.Drawing.Color fillColor)
    {
        // Create bitmap
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            // Obtain graphics object
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                // Fill background
                graphics.Clear(fillColor);
            }

            // Save to file (extension determines format)
            bitmap.Save(filePath);
        }
    }

    // Creates a DOCX file that contains a single image and some text
    private static void CreateSampleDocumentWithImage(string docPath, string imagePath, string title)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln(title);
        builder.InsertParagraph();
        builder.InsertImage(imagePath);
        builder.InsertParagraph();

        doc.Save(docPath);
    }

    // Generates a thumbnail from an existing image file
    private static void CreateThumbnail(string sourcePath, string thumbPath, int maxWidth, int maxHeight)
    {
        // Load source image
        using (Bitmap sourceBitmap = new Bitmap(sourcePath))
        {
            // Determine scaling factor while preserving aspect ratio
            double ratioX = (double)maxWidth / sourceBitmap.Width;
            double ratioY = (double)maxHeight / sourceBitmap.Height;
            double ratio = Math.Min(ratioX, ratioY);
            int thumbWidth = (int)(sourceBitmap.Width * ratio);
            int thumbHeight = (int)(sourceBitmap.Height * ratio);

            // Create thumbnail bitmap
            using (Bitmap thumbBitmap = new Bitmap(thumbWidth, thumbHeight))
            {
                using (Graphics graphics = Graphics.FromImage(thumbBitmap))
                {
                    // High quality scaling (InterpolationMode omitted to avoid missing enum)
                    graphics.DrawImage(sourceBitmap, 0, 0, thumbWidth, thumbHeight);
                }

                // Save thumbnail
                thumbBitmap.Save(thumbPath);
            }
        }
    }
}
