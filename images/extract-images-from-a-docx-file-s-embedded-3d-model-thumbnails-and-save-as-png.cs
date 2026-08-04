using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Folder for all generated files.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // 1. Create a sample PNG image that will act as a 3D model thumbnail.
        string sampleImagePath = Path.Combine(artifactsDir, "sample.png");
        CreateSampleImage(sampleImagePath, 200, 200);

        // 2. Create a DOCX file and insert the sample image.
        string docPath = Path.Combine(artifactsDir, "sample.docx");
        CreateDocumentWithImage(docPath, sampleImagePath);

        // 3. Extract all images from the DOCX and save them as PNG files.
        ExtractImagesAsPng(docPath, artifactsDir);
    }

    // Creates a deterministic PNG image using Aspose.Drawing.
    private static void CreateSampleImage(string path, int width, int height)
    {
        var bitmap = new Bitmap(width, height);
        var graphics = Graphics.FromImage(bitmap);
        graphics.Clear(Color.LightBlue);
        graphics.DrawRectangle(new Pen(Color.DarkBlue, 5), 10, 10, width - 20, height - 20);
        graphics.Dispose();
        bitmap.Save(path);
        bitmap.Dispose();
    }

    // Creates a DOCX document and inserts the provided image.
    private static void CreateDocumentWithImage(string docPath, string imagePath)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath);
    }

    // Extracts every image from the document and saves it as a PNG file.
    private static void ExtractImagesAsPng(string docPath, string outputDir)
    {
        var doc = new Document(docPath);
        var shapes = doc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;

        foreach (Shape shape in shapes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                // Retrieve the raw image bytes from the shape.
                byte[] imageBytes = shape.ImageData.ToByteArray();

                // Load the bytes into an Aspose.Drawing bitmap and save as PNG.
                using (var ms = new MemoryStream(imageBytes))
                {
                    var bitmap = new Bitmap(ms);
                    string outPath = Path.Combine(outputDir, $"thumb_{imageIndex}.png");
                    bitmap.Save(outPath);
                    bitmap.Dispose();
                }

                imageIndex++;
            }
        }

        if (imageIndex == 0)
            throw new InvalidOperationException("No images were extracted from the document.");
    }
}
