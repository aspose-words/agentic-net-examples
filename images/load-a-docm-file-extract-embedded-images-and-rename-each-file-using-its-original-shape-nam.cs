using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Prepare folders
        string baseDir = Directory.GetCurrentDirectory();
        string artifactsDir = Path.Combine(baseDir, "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // 1. Create a deterministic sample image (sample.png)
        string sampleImagePath = Path.Combine(artifactsDir, "sample.png");
        CreateSampleImage(sampleImagePath, 200, 200);

        // 2. Build a DOCM document and embed the image with a custom shape name
        string docmPath = Path.Combine(artifactsDir, "sample.docm");
        CreateDocumentWithImage(docmPath, sampleImagePath);

        // 3. Load the DOCM document
        Document doc = new Document(docmPath);

        // 4. Extract each embedded image and rename the file using the shape's original name
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage) continue;

            // Ensure the shape has a name; if not, generate a fallback
            string shapeName = !string.IsNullOrEmpty(shape.Name) ? shape.Name : $"Shape_{extractedCount}";
            string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
            string outputFile = Path.Combine(artifactsDir, $"{shapeName}{extension}");

            // Save the image
            shape.ImageData.Save(outputFile);
            extractedCount++;
        }

        // Validation: at least one image must have been extracted
        if (extractedCount == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        // Program ends automatically
    }

    // Creates a simple PNG image using Aspose.Drawing
    private static void CreateSampleImage(string filePath, int width, int height)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.White);
                // Draw a simple rectangle
                graphics.FillRectangle(new SolidBrush(Color.LightBlue), 10, 10, width - 20, height - 20);
            }
            bitmap.Save(filePath, ImageFormat.Png);
        }
    }

    // Creates a DOCM file, inserts the image, and assigns a custom name to the shape
    private static void CreateDocumentWithImage(string docmPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the image; the returned Shape represents the picture
        Shape pictureShape = builder.InsertImage(imagePath);
        pictureShape.Name = "EmbeddedSampleImage"; // Custom shape name for later renaming

        // Save as a macro-enabled document (DOCM)
        doc.Save(docmPath, SaveFormat.Docm);
    }
}
