using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Base directories
        string baseDir = AppDomain.CurrentDomain.BaseDirectory;
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string outputDir = Path.Combine(baseDir, "ExtractedImages");
        string sampleImagePath = Path.Combine(baseDir, "sample.png");

        // Ensure directories exist
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create a deterministic sample image
        CreateSampleImage(sampleImagePath);

        // Create sample ODT documents containing the image
        CreateSampleDocuments(inputDir, sampleImagePath, 2);

        // Batch extract images from ODT files
        foreach (string docPath in Directory.GetFiles(inputDir, "*.odt"))
        {
            // Load the document
            Document doc = new Document(docPath);
            string docName = Path.GetFileNameWithoutExtension(docPath);
            string docOutputDir = Path.Combine(outputDir, docName);
            Directory.CreateDirectory(docOutputDir);

            // Extract images from Shape nodes
            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;

            foreach (Shape shape in shapes)
            {
                if (shape.HasImage)
                {
                    imageIndex++;
                    // Determine image extension based on image type
                    string ext = shape.ImageData.ImageType.ToString().ToLower();
                    string outPath = Path.Combine(docOutputDir, $"image{imageIndex}.{ext}");
                    shape.ImageData.Save(outPath);
                }
            }

            // Validate that at least one image was extracted
            if (imageIndex == 0)
                throw new Exception($"No images were extracted from document '{docName}'.");
        }

        // Indicate completion
        Console.WriteLine("Image extraction completed successfully.");
    }

    private static void CreateSampleImage(string path)
    {
        // Create a 100x100 pixel bitmap with a solid color
        Bitmap bitmap = new Bitmap(100, 100);
        Graphics graphics = Graphics.FromImage(bitmap);
        graphics.Clear(Color.LightBlue);
        // Save the bitmap to a PNG file
        bitmap.Save(path);
        // Clean up resources
        graphics.Dispose();
        bitmap.Dispose();
    }

    private static void CreateSampleDocuments(string folder, string imagePath, int count)
    {
        for (int i = 1; i <= count; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"Sample Document {i}");
            // Insert the previously created sample image
            builder.InsertImage(imagePath);
            string docPath = Path.Combine(folder, $"doc{i}.odt");
            doc.Save(docPath);
        }
    }
}
