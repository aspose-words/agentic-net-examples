using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class BatchImageExtractor
{
    public static void Main()
    {
        // Prepare folders
        string baseDir = Directory.GetCurrentDirectory();
        string inputFolder = Path.Combine(baseDir, "InputDocs");
        string outputFolder = Path.Combine(baseDir, "ExtractedImages");
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create a deterministic sample image (input.png)
        string sampleImagePath = Path.Combine(baseDir, "input.png");
        CreateSampleImage(sampleImagePath);

        // Create sample ODT documents containing the image
        CreateSampleOdtDocuments(inputFolder, sampleImagePath, 2);

        // Process each ODT file: extract images and organize by document name
        foreach (string odtPath in Directory.GetFiles(inputFolder, "*.odt"))
        {
            // Load the document
            Document doc = new Document(odtPath);

            // Prepare output subfolder named after the original document (without extension)
            string docName = Path.GetFileNameWithoutExtension(odtPath);
            string docOutputFolder = Path.Combine(outputFolder, docName);
            Directory.CreateDirectory(docOutputFolder);

            // Extract images
            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;
            foreach (Shape shape in shapes)
            {
                if (shape.HasImage)
                {
                    string imagePath = Path.Combine(docOutputFolder, $"image-{imageIndex}.png");
                    shape.ImageData.Save(imagePath);
                    imageIndex++;
                }
            }

            // Validate that at least one image was extracted
            if (imageIndex == 0)
                throw new InvalidOperationException($"No images were extracted from document '{odtPath}'.");
        }

        Console.WriteLine("Image extraction completed successfully.");
    }

    // Creates a simple white PNG image using Aspose.Drawing
    private static void CreateSampleImage(string filePath)
    {
        const int width = 100;
        const int height = 100;
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.White);
                // Draw a simple black rectangle for visual distinction
                using (Pen pen = new Pen(Color.Black, 2))
                {
                    graphics.DrawRectangle(pen, 10, 10, width - 20, height - 20);
                }
            }
            bitmap.Save(filePath, ImageFormat.Png);
        }

        // Validate that the image file was created
        if (!File.Exists(filePath))
            throw new InvalidOperationException($"Failed to create sample image at '{filePath}'.");
    }

    // Generates a given number of ODT documents each containing the sample image
    private static void CreateSampleOdtDocuments(string folderPath, string imagePath, int count)
    {
        for (int i = 1; i <= count; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"Document {i}");
            builder.InsertImage(imagePath);
            string docPath = Path.Combine(folderPath, $"Doc{i}.odt");
            doc.Save(docPath, SaveFormat.Odt);

            // Validate that the document was saved
            if (!File.Exists(docPath))
                throw new InvalidOperationException($"Failed to save ODT document at '{docPath}'.");
        }
    }
}
