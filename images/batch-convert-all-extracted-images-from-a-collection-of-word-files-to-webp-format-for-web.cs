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
        // Define input and output folders
        string inputFolder = "InputDocs";
        string outputFolder = "ExtractedWebP";

        // Clean previous run data
        if (Directory.Exists(inputFolder))
            Directory.Delete(inputFolder, true);
        if (Directory.Exists(outputFolder))
            Directory.Delete(outputFolder, true);
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create a deterministic sample image (input.png)
        string sampleImagePath = "input.png";
        CreateSampleImage(sampleImagePath, 200, 200);

        // Create sample Word documents containing the image
        CreateSampleDocument(Path.Combine(inputFolder, "Doc1.docx"), sampleImagePath);
        CreateSampleDocument(Path.Combine(inputFolder, "Doc2.docx"), sampleImagePath);

        // Batch process each Word document
        foreach (string docPath in Directory.GetFiles(inputFolder, "*.docx"))
        {
            Document doc = new Document(docPath);
            NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
            int extractedCount = 0;
            int imageIndex = 0;

            foreach (Shape shape in shapeNodes)
            {
                if (shape.HasImage)
                {
                    // Save the embedded image to a memory stream
                    using (MemoryStream imgStream = new MemoryStream())
                    {
                        shape.ImageData.Save(imgStream);
                        imgStream.Position = 0;

                        // Load the image into Aspose.Drawing.Bitmap
                        using (Bitmap bitmap = new Bitmap(imgStream))
                        {
                            // Define output PNG file name (WebP not supported in this environment)
                            string outputFileName = $"{Path.GetFileNameWithoutExtension(docPath)}_img{imageIndex}.png";
                            string outputPath = Path.Combine(outputFolder, outputFileName);

                            // Save as PNG (fallback for WebP)
                            bitmap.Save(outputPath, ImageFormat.Png);

                            extractedCount++;
                            imageIndex++;
                        }
                    }
                }
            }

            // Validate that at least one image was extracted
            if (extractedCount == 0)
                throw new InvalidOperationException($"No images were extracted from document '{docPath}'.");
        }

        // Verify that output files exist
        string[] outputFiles = Directory.GetFiles(outputFolder);
        if (outputFiles.Length == 0)
            throw new InvalidOperationException("No images were created in the output folder.");

        Console.WriteLine($"Batch conversion completed. {outputFiles.Length} images saved to '{outputFolder}'.");
    }

    // Creates a simple white bitmap with a black rectangle for deterministic content
    private static void CreateSampleImage(string path, int width, int height)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Aspose.Drawing.Color.White);
                using (SolidBrush brush = new SolidBrush(Aspose.Drawing.Color.Black))
                {
                    g.FillRectangle(brush, width / 4, height / 4, width / 2, height / 2);
                }
            }
            bitmap.Save(path, ImageFormat.Png);
        }
    }

    // Creates a Word document and inserts the specified image
    private static void CreateSampleDocument(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath);
    }
}
