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
        // Prepare output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // ---------- Create sample images ----------
        string imgPath1 = Path.Combine(artifactsDir, "sample1.png");
        string imgPath2 = Path.Combine(artifactsDir, "sample2.png");

        CreateSampleImage(imgPath1, 100, 100, Aspose.Drawing.Color.LightBlue);
        CreateSampleImage(imgPath2, 120, 80, Aspose.Drawing.Color.LightCoral);

        // ---------- Build a DOCM with the images ----------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert first image and give the shape a name.
        Shape shape1 = builder.InsertImage(imgPath1);
        shape1.Name = "FirstImage";

        // Insert second image and give the shape a name.
        builder.InsertParagraph();
        Shape shape2 = builder.InsertImage(imgPath2);
        shape2.Name = "SecondImage";

        // Save as DOCM.
        string docPath = Path.Combine(artifactsDir, "Sample.docm");
        doc.Save(docPath, SaveFormat.Docm);

        // ---------- Load the DOCM and extract images ----------
        Document loadedDoc = new Document(docPath);
        var shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true)
                                  .OfType<Shape>()
                                  .Where(s => s.HasImage)
                                  .ToList();

        int extractedCount = 0;
        foreach (var shape in shapeNodes)
        {
            // Determine a safe file name based on the shape's original name.
            string baseName = string.IsNullOrWhiteSpace(shape.Name) ? $"shape_{extractedCount}" : shape.Name;
            string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
            string outFile = Path.Combine(artifactsDir, $"{baseName}{extension}");

            // Save the image.
            shape.ImageData.Save(outFile);
            extractedCount++;
        }

        // Validate that at least one image was extracted.
        if (extractedCount == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        // Optional: indicate completion.
        Console.WriteLine($"Extracted {extractedCount} image(s) to \"{artifactsDir}\".");
    }

    // Helper to create a deterministic bitmap using Aspose.Drawing.
    private static void CreateSampleImage(string filePath, int width, int height, Aspose.Drawing.Color fillColor)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(fillColor);
            }
            bitmap.Save(filePath, ImageFormat.Png);
        }
    }
}
