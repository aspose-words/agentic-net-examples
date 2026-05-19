using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Loading;
using Aspose.Words.Saving;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare output directories.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        string imageDir = Path.Combine(baseDir, "Images");
        Directory.CreateDirectory(imageDir);
        Directory.CreateDirectory(baseDir);

        // -----------------------------------------------------------------
        // 1. Create sample images using Aspose.Drawing.
        // -----------------------------------------------------------------
        string imgPath1 = Path.Combine(imageDir, "sample1.png");
        string imgPath2 = Path.Combine(imageDir, "sample2.png");

        CreateSampleImage(imgPath1, 200, 200, Aspose.Drawing.Color.LightBlue);
        CreateSampleImage(imgPath2, 150, 150, Aspose.Drawing.Color.LightCoral);

        // -----------------------------------------------------------------
        // 2. Build a DOCM file and embed the images as shapes with names.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert first image and assign a shape name.
        Shape shape1 = builder.InsertImage(imgPath1);
        shape1.Name = "FirstImage";

        // Insert a paragraph break between images.
        builder.InsertParagraph();

        // Insert second image and assign a shape name.
        Shape shape2 = builder.InsertImage(imgPath2);
        shape2.Name = "SecondImage";

        // Save the document as a macro-enabled DOCM file.
        string docPath = Path.Combine(baseDir, "SampleDocument.docm");
        doc.Save(docPath, SaveFormat.Docm);

        // -----------------------------------------------------------------
        // 3. Load the DOCM file and extract embedded images.
        // -----------------------------------------------------------------
        LoadOptions loadOptions = new LoadOptions(); // default options
        Document loadedDoc = new Document(docPath, loadOptions);

        var shapeImages = loadedDoc
            .GetChildNodes(NodeType.Shape, true)
            .OfType<Shape>()
            .Where(s => s.HasImage)
            .ToList();

        if (!shapeImages.Any())
            throw new InvalidOperationException("No images were found in the document.");

        int index = 0;
        foreach (var shape in shapeImages)
        {
            // Determine file extension based on the image type stored in the shape.
            string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);

            // Use the shape's Name property; fall back to an index if the name is empty.
            string baseFileName = !string.IsNullOrEmpty(shape.Name)
                ? shape.Name
                : $"Image_{index}";

            string outputImagePath = Path.Combine(imageDir, $"{baseFileName}{extension}");

            // Save the image to the file system.
            shape.ImageData.Save(outputImagePath);

            index++;
        }

        // -----------------------------------------------------------------
        // 4. Validation – ensure at least one image file was written.
        // -----------------------------------------------------------------
        var savedFiles = Directory.GetFiles(imageDir, "*.*")
            .Where(f => f.EndsWith(".png", StringComparison.OrdinalIgnoreCase) ||
                        f.EndsWith(".jpg", StringComparison.OrdinalIgnoreCase) ||
                        f.EndsWith(".jpeg", StringComparison.OrdinalIgnoreCase) ||
                        f.EndsWith(".bmp", StringComparison.OrdinalIgnoreCase) ||
                        f.EndsWith(".gif", StringComparison.OrdinalIgnoreCase) ||
                        f.EndsWith(".tiff", StringComparison.OrdinalIgnoreCase) ||
                        f.EndsWith(".webp", StringComparison.OrdinalIgnoreCase))
            .ToArray();

        if (savedFiles.Length == 0)
            throw new InvalidOperationException("Image extraction failed – no image files were created.");

        // Program completed successfully.
    }

    // Helper method to create a deterministic sample image.
    private static void CreateSampleImage(string filePath, int width, int height, Aspose.Drawing.Color backgroundColor)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(backgroundColor);
            // Additional deterministic drawing can be added here if needed.
            bitmap.Save(filePath);
        }
    }
}
