using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;
using Aspose.Drawing; // Aspose.Drawing.Common namespace

public class Program
{
    public static void Main()
    {
        // Define base folder for all generated files.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // 1. Create a sample image that will be inserted into the source DOCX.
        // -----------------------------------------------------------------
        string sampleImagePath = Path.Combine(artifactsDir, "sample.png");
        CreateSampleImage(sampleImagePath, 200, 150, Aspose.Drawing.Color.LightBlue);

        // -----------------------------------------------------------------
        // 2. Build a sample DOCX containing the image.
        // -----------------------------------------------------------------
        string sourceDocPath = Path.Combine(artifactsDir, "input.docx");
        CreateDocumentWithImage(sourceDocPath, sampleImagePath);

        // -----------------------------------------------------------------
        // 3. Load the DOCX, replace every image with a placeholder image.
        // -----------------------------------------------------------------
        Document doc = new Document(sourceDocPath);

        // Create placeholder image (a gray box).
        string placeholderImagePath = Path.Combine(artifactsDir, "placeholder.png");
        CreateSampleImage(placeholderImagePath, 100, 100, Aspose.Drawing.Color.LightGray);

        // Replace each shape that has an image.
        var shapes = doc.GetChildNodes(NodeType.Shape, true)
                        .Cast<Shape>()
                        .Where(s => s.HasImage);
        foreach (Shape shape in shapes)
        {
            shape.ImageData.SetImage(placeholderImagePath);
        }

        // -----------------------------------------------------------------
        // 4. Export the modified document to HTML, preserving layout.
        // -----------------------------------------------------------------
        string htmlOutputPath = Path.Combine(artifactsDir, "output.html");
        string imagesFolder = Path.Combine(artifactsDir, "HtmlImages");
        Directory.CreateDirectory(imagesFolder);

        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            ImagesFolder = imagesFolder,
            ExportImagesAsBase64 = false, // keep images as separate files
            ScaleImageToShapeSize = true   // preserve layout scaling
        };

        doc.Save(htmlOutputPath, htmlOptions);

        // -----------------------------------------------------------------
        // 5. Validate that output files were created.
        // -----------------------------------------------------------------
        if (!File.Exists(htmlOutputPath))
            throw new InvalidOperationException("HTML file was not created.");

        var exportedImages = Directory.GetFiles(imagesFolder);
        if (exportedImages.Length == 0)
            throw new InvalidOperationException("No images were exported during HTML conversion.");

        // The example finishes without requiring user interaction.
    }

    // Helper to create a simple solid‑color PNG image.
    private static void CreateSampleImage(string filePath, int width, int height, Aspose.Drawing.Color backColor)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(backColor);
            }
            bitmap.Save(filePath);
        }
    }

    // Helper to create a DOCX that contains a single image.
    private static void CreateDocumentWithImage(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Document with an image:");
        builder.InsertImage(imagePath);
        builder.Writeln();
        builder.Writeln("End of document.");

        doc.Save(docPath);
    }
}
