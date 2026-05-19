using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;
using Aspose.Drawing; // Aspose.Drawing.Common provides Bitmap, Graphics, Color, Font

public class Program
{
    public static void Main()
    {
        // Prepare folders
        string baseDir = Directory.GetCurrentDirectory();
        string artifactsDir = Path.Combine(baseDir, "Artifacts");
        string imagesInputDir = Path.Combine(artifactsDir, "InputImages");
        string imagesOutputDir = Path.Combine(artifactsDir, "HtmlImages");
        Directory.CreateDirectory(artifactsDir);
        Directory.CreateDirectory(imagesInputDir);
        Directory.CreateDirectory(imagesOutputDir);

        // 1. Create a sample image that will be inserted into the document
        string sampleImagePath = Path.Combine(imagesInputDir, "sample.png");
        CreateSampleImage(sampleImagePath, 100, 100, Aspose.Drawing.Color.Blue, "Sample");

        // 2. Build a DOCX containing several copies of the sample image
        string docxPath = Path.Combine(artifactsDir, "input.docx");
        CreateDocumentWithImages(docxPath, sampleImagePath);

        // 3. Load the document
        Document doc = new Document(docxPath);

        // 4. Create a placeholder image that will replace every existing image
        string placeholderImagePath = Path.Combine(imagesInputDir, "placeholder.png");
        CreateSampleImage(placeholderImagePath, 100, 100, Aspose.Drawing.Color.LightGray, "Placeholder");

        // 5. Replace each image in the document with the placeholder image
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                // Replace the image data with the placeholder image file
                shape.ImageData.SetImage(placeholderImagePath);
            }
        }

        // 6. Save the modified document as HTML, preserving layout
        string htmlPath = Path.Combine(artifactsDir, "output.html");
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            ImagesFolder = imagesOutputDir,
            ExportImagesAsBase64 = false // keep images as separate files
        };
        doc.Save(htmlPath, htmlOptions);

        // 7. Simple validation
        if (!File.Exists(htmlPath))
            throw new InvalidOperationException("HTML file was not created.");

        string[] exportedImages = Directory.GetFiles(imagesOutputDir);
        if (exportedImages.Length == 0)
            throw new InvalidOperationException("No images were exported during HTML conversion.");

        // The example finishes without requiring user interaction.
    }

    // Creates a simple bitmap with a solid background and centered text.
    private static void CreateSampleImage(string filePath, int width, int height, Aspose.Drawing.Color backColor, string text)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(backColor);
                using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 12))
                {
                    // Measure the text to center it
                    var textSize = graphics.MeasureString(text, font);
                    float x = (width - textSize.Width) / 2f;
                    float y = (height - textSize.Height) / 2f;
                    graphics.DrawString(text, font, Aspose.Drawing.Brushes.Black, x, y);
                }
            }
            bitmap.Save(filePath);
        }
    }

    // Builds a document and inserts the provided image multiple times.
    private static void CreateDocumentWithImages(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Document with images. Each image will be replaced by a placeholder.");
        for (int i = 0; i < 3; i++)
        {
            builder.InsertImage(imagePath);
            builder.Writeln(); // add a line break after each image
        }

        doc.Save(docPath);
    }
}
