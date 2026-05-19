using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;

public class BatchImageExtractor
{
    public static void Main()
    {
        // Base working directory.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        string inputDir = Path.Combine(baseDir, "Input");
        string outputDir = Path.Combine(baseDir, "Output");

        // Prepare folders.
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create a deterministic sample image.
        string sampleImagePath = Path.Combine(baseDir, "sample.png");
        CreateSampleImage(sampleImagePath, 200, 200);

        // Create sample ODT documents containing the image.
        CreateSampleDocument(Path.Combine(inputDir, "DocumentA.odt"), sampleImagePath);
        CreateSampleDocument(Path.Combine(inputDir, "DocumentB.odt"), sampleImagePath);

        // Batch extract images from all ODT files.
        foreach (string odtPath in Directory.GetFiles(inputDir, "*.odt"))
        {
            // Load the ODT document.
            Document doc = new Document(odtPath);

            // Collect all shapes that contain images.
            var imageShapes = doc.GetChildNodes(NodeType.Shape, true)
                                 .Cast<Shape>()
                                 .Where(s => s.HasImage)
                                 .ToList();

            if (imageShapes.Count == 0)
                throw new InvalidOperationException($"No images found in document '{odtPath}'.");

            // Create a subfolder named after the source document (without extension).
            string docName = Path.GetFileNameWithoutExtension(odtPath);
            string docOutputDir = Path.Combine(outputDir, docName);
            Directory.CreateDirectory(docOutputDir);

            int imageIndex = 0;
            foreach (Shape shape in imageShapes)
            {
                // Determine proper file extension for the image type.
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string imageFileName = $"image_{imageIndex}{extension}";
                string imagePath = Path.Combine(docOutputDir, imageFileName);

                // Save the image to the file system.
                shape.ImageData.Save(imagePath);
                imageIndex++;
            }

            // Validate that at least one image file was written.
            if (Directory.GetFiles(docOutputDir).Length == 0)
                throw new InvalidOperationException($"Failed to extract images for document '{odtPath}'.");
        }
    }

    // Creates a simple white bitmap with a black rectangle and saves it to the given path.
    private static void CreateSampleImage(string filePath, int width, int height)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(Color.White);
            // Draw a simple black rectangle.
            graphics.DrawRectangle(new Pen(Color.Black, 5), 10, 10, width - 20, height - 20);
            bitmap.Save(filePath);
        }
    }

    // Creates an ODT document with a single image inserted.
    private static void CreateSampleDocument(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln($"This is a sample document: {Path.GetFileName(docPath)}");
        builder.InsertImage(imagePath);
        doc.Save(docPath, SaveFormat.Odt);
    }
}
