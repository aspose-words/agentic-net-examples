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
        // Prepare folders
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        string inputFolder = Path.Combine(baseDir, "InputDocs");
        string outputFolder = Path.Combine(baseDir, "ExtractedImages");
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create a deterministic sample image
        string sampleImagePath = Path.Combine(baseDir, "sample.png");
        CreateSampleImage(sampleImagePath);

        // Create a few ODT documents that contain the sample image
        for (int i = 1; i <= 2; i++)
        {
            string docPath = Path.Combine(inputFolder, $"Document{i}.odt");
            CreateOdtWithImage(docPath, sampleImagePath);
        }

        // Batch extract images from each ODT file
        foreach (string odtFile in Directory.GetFiles(inputFolder, "*.odt"))
        {
            string docName = Path.GetFileNameWithoutExtension(odtFile);
            string docOutputFolder = Path.Combine(outputFolder, docName);
            Directory.CreateDirectory(docOutputFolder);

            Document doc = new Document(odtFile);
            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

            int imageIndex = 0;
            foreach (Shape shape in shapes.OfType<Shape>())
            {
                if (shape.HasImage)
                {
                    string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                    string outPath = Path.Combine(docOutputFolder, $"Image{imageIndex}{extension}");
                    shape.ImageData.Save(outPath);
                    imageIndex++;
                }
            }

            if (imageIndex == 0)
                throw new InvalidOperationException($"No images were extracted from '{odtFile}'.");
        }

        Console.WriteLine("Image extraction completed.");
    }

    private static void CreateSampleImage(string filePath)
    {
        // 100x100 white bitmap
        using (Bitmap bitmap = new Bitmap(100, 100))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(Color.White);
            // Draw a simple black rectangle for visual distinction
            graphics.DrawRectangle(new Pen(Color.Black), 10, 10, 80, 80);
            bitmap.Save(filePath);
        }
    }

    private static void CreateOdtWithImage(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        // Save as ODT (OpenDocument Text)
        doc.Save(docPath, SaveFormat.Odt);
    }
}
