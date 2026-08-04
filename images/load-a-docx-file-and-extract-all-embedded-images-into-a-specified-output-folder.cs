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
        // Define folders.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "ImageExample");
        string inputDir = Path.Combine(baseDir, "Input");
        string outputDir = Path.Combine(baseDir, "Output");
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create a deterministic sample image.
        string sampleImagePath = Path.Combine(inputDir, "sample.png");
        using (Bitmap bitmap = new Bitmap(200, 200))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(Color.White);
            // Simple drawing – a black rectangle.
            graphics.DrawRectangle(new Pen(Color.Black), 20, 20, 160, 160);
            bitmap.Save(sampleImagePath);
        }

        // Create a DOCX document and insert the sample image.
        string docPath = Path.Combine(inputDir, "sample.docx");
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        doc.Save(docPath);

        // Load the document.
        Document loadedDoc = new Document(docPath);

        // Extract all images from the document.
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string outFile = Path.Combine(outputDir, $"extracted_{imageIndex}{extension}");
                shape.ImageData.Save(outFile);
                imageIndex++;
            }
        }

        // Validate that at least one image was extracted.
        if (imageIndex == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        // Optional: indicate completion.
        Console.WriteLine($"Extracted {imageIndex} image(s) to \"{outputDir}\".");
    }
}
