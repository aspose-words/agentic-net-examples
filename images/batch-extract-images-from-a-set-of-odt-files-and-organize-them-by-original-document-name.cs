using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Set up folders.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        string inputDir = Path.Combine(baseDir, "Input");
        string outputDir = Path.Combine(baseDir, "Output");
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create a deterministic sample image.
        string sampleImagePath = Path.Combine(inputDir, "sample.png");
        CreateSampleImage(sampleImagePath, 200, 200);

        // Create sample ODT documents that contain the image.
        CreateSampleDocument(Path.Combine(inputDir, "Doc1.odt"), sampleImagePath, 2);
        CreateSampleDocument(Path.Combine(inputDir, "Doc2.odt"), sampleImagePath, 3);

        // Batch extract images from all ODT files.
        foreach (string odtPath in Directory.GetFiles(inputDir, "*.odt"))
        {
            // Load the document.
            Document doc = new Document(odtPath);

            // Collect all shapes that have images.
            var imageShapes = doc.GetChildNodes(NodeType.Shape, true)
                                 .Cast<Shape>()
                                 .Where(s => s.HasImage)
                                 .ToList();

            if (imageShapes.Count == 0)
                throw new InvalidOperationException($"No images found in document '{odtPath}'.");

            // Prepare output folder for this document.
            string docName = Path.GetFileNameWithoutExtension(odtPath);
            string docOutputDir = Path.Combine(outputDir, docName);
            Directory.CreateDirectory(docOutputDir);

            int imageIndex = 0;
            foreach (Shape shape in imageShapes)
            {
                // Determine file extension based on image type.
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string outFile = Path.Combine(docOutputDir, $"Image_{imageIndex}{extension}");

                // Save the image.
                shape.ImageData.Save(outFile);
                imageIndex++;
            }
        }
    }

    // Creates a simple white PNG image with a colored rectangle.
    private static void CreateSampleImage(string filePath, int width, int height)
    {
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height);
        Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap);
        graphics.Clear(Aspose.Drawing.Color.White);
        using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Blue, 5))
        {
            graphics.DrawRectangle(pen, 10, 10, width - 20, height - 20);
        }
        bitmap.Save(filePath);
        graphics.Dispose();
        bitmap.Dispose();
    }

    // Creates an ODT document with the specified number of inserted images.
    private static void CreateSampleDocument(string docPath, string imagePath, int imageCount)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        for (int i = 0; i < imageCount; i++)
        {
            builder.Writeln($"Image #{i + 1}:");
            builder.InsertImage(imagePath);
            builder.Writeln();
        }

        // Save as ODT.
        doc.Save(docPath, SaveFormat.Odt);
    }
}
