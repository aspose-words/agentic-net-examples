using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Words.Loading;
using Aspose.Drawing; // Aspose.Drawing.Common provides Bitmap, Graphics, Color

public class Program
{
    public static void Main()
    {
        // Define working folders
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "ImageCatalogDemo");
        string inputDocsDir = Path.Combine(baseDir, "InputDocs");
        string extractedImagesDir = Path.Combine(baseDir, "ExtractedImages");
        string thumbnailsDir = Path.Combine(baseDir, "Thumbnails");
        string outputDir = Path.Combine(baseDir, "Output");

        // Ensure folders exist
        Directory.CreateDirectory(inputDocsDir);
        Directory.CreateDirectory(extractedImagesDir);
        Directory.CreateDirectory(thumbnailsDir);
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a deterministic sample image that will be used in the docs
        // -----------------------------------------------------------------
        string sampleImagePath = Path.Combine(baseDir, "sample.png");
        CreateSampleImage(sampleImagePath, 200, 200, Aspose.Drawing.Color.LightBlue, Aspose.Drawing.Color.DarkBlue);

        // -----------------------------------------------------------------
        // 2. Generate a few sample DOCX files that contain the sample image
        // -----------------------------------------------------------------
        const int docCount = 3;
        for (int i = 1; i <= docCount; i++)
        {
            string docPath = Path.Combine(inputDocsDir, $"Document{i}.docx");
            CreateDocumentWithImage(docPath, sampleImagePath, $"Sample document {i}");
        }

        // -----------------------------------------------------------------
        // 3. Process each DOCX: extract images, create thumbnails
        // -----------------------------------------------------------------
        var docFiles = Directory.GetFiles(inputDocsDir, "*.docx");
        int globalImageIndex = 0;

        foreach (var docFile in docFiles)
        {
            // Load the source document
            Document srcDoc = new Document(docFile);

            // Get all shapes that contain images
            var shapeNodes = srcDoc.GetChildNodes(NodeType.Shape, true)
                                   .Cast<Shape>()
                                   .Where(s => s.HasImage);

            foreach (var shape in shapeNodes)
            {
                // Save the original image
                string imageExtension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string imageFileName = $"Img_{globalImageIndex}{imageExtension}";
                string imageFullPath = Path.Combine(extractedImagesDir, imageFileName);
                shape.ImageData.Save(imageFullPath);
                globalImageIndex++;

                // Create a thumbnail for the saved image
                string thumbFileName = Path.GetFileNameWithoutExtension(imageFileName) + "_thumb.png";
                string thumbFullPath = Path.Combine(thumbnailsDir, thumbFileName);
                CreateThumbnailFromImage(imageFullPath, thumbFullPath, 100, 100);
            }
        }

        // -----------------------------------------------------------------
        // 4. Build a PDF catalog that shows all thumbnails
        // -----------------------------------------------------------------
        Document catalog = new Document();
        DocumentBuilder builder = new DocumentBuilder(catalog);

        builder.Writeln("Image Catalog");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln();

        var thumbnailFiles = Directory.GetFiles(thumbnailsDir, "*_thumb.png");
        foreach (var thumbPath in thumbnailFiles)
        {
            // Insert thumbnail image
            builder.InsertImage(thumbPath);
            builder.Writeln(); // add spacing
        }

        // Save the final PDF catalog
        string catalogPath = Path.Combine(outputDir, "ImageCatalog.pdf");
        catalog.Save(catalogPath, SaveFormat.Pdf);

        Console.WriteLine($"Catalog created at: {catalogPath}");
    }

    // Creates a simple PNG image using Aspose.Drawing
    private static void CreateSampleImage(string filePath, int width, int height, Aspose.Drawing.Color background, Aspose.Drawing.Color rectangle)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(background);
                // Draw a simple rectangle
                int rectSize = Math.Min(width, height) / 2;
                int offset = (width - rectSize) / 2;
                g.FillRectangle(new SolidBrush(rectangle), offset, offset, rectSize, rectSize);
            }
            bitmap.Save(filePath);
        }
    }

    // Creates a DOCX with a single image and saves it
    private static void CreateDocumentWithImage(string docPath, string imagePath, string title)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln(title);
        builder.InsertImage(imagePath);
        doc.Save(docPath);
    }

    // Generates a thumbnail by inserting the image into a temporary document and rendering it at a lower resolution
    private static void CreateThumbnailFromImage(string sourceImagePath, string thumbnailPath, int thumbWidth, int thumbHeight)
    {
        // Load image into a temporary document
        Document tempDoc = new Document();
        DocumentBuilder tempBuilder = new DocumentBuilder(tempDoc);
        Shape imgShape = tempBuilder.InsertImage(sourceImagePath);
        imgShape.Width = thumbWidth;
        imgShape.Height = thumbHeight;

        // Render the document page (which contains only the image) to a PNG thumbnail
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png)
        {
            Resolution = 72,
            ImageSize = new System.Drawing.Size(thumbWidth, thumbHeight) // System.Drawing.Size is allowed here as part of Aspose.Words options
        };
        tempDoc.Save(thumbnailPath, options);
    }
}
