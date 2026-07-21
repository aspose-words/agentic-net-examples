using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Prepare folders
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string outputDir = Path.Combine(baseDir, "OutputImages");
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create sample image to be inserted into documents
        string sampleImagePath = Path.Combine(baseDir, "sample.png");
        CreateSamplePng(sampleImagePath, 200, 200);

        // Create a few sample Word documents containing the image
        for (int docIdx = 0; docIdx < 2; docIdx++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"Document {docIdx + 1}");
            builder.InsertImage(sampleImagePath);
            string docPath = Path.Combine(inputDir, $"Doc{docIdx + 1}.docx");
            doc.Save(docPath);
        }

        // Batch process: extract images and convert each to WebP
        int totalConverted = 0;
        var docFiles = Directory.GetFiles(inputDir, "*.docx");
        foreach (var docFile in docFiles)
        {
            Document doc = new Document(docFile);
            NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;
            foreach (Shape shape in shapeNodes.OfType<Shape>())
            {
                if (!shape.HasImage) continue;

                // Save the shape's image to a memory stream
                using (MemoryStream imgStream = new MemoryStream())
                {
                    shape.ImageData.Save(imgStream);
                    imgStream.Position = 0;

                    // Create a temporary document that contains only this image
                    Document tempDoc = new Document();
                    DocumentBuilder tempBuilder = new DocumentBuilder(tempDoc);
                    tempBuilder.InsertImage(imgStream);

                    // Define WebP save options
                    ImageSaveOptions webpOptions = new ImageSaveOptions(SaveFormat.WebP)
                    {
                        // Optional: set resolution or other options if needed
                        Resolution = 96
                    };

                    // Build output file name
                    string outFileName = $"converted_{Path.GetFileNameWithoutExtension(docFile)}_{imageIndex}.webp";
                    string outPath = Path.Combine(outputDir, outFileName);

                    // Render the temporary document (single page) to WebP
                    tempDoc.Save(outPath, webpOptions);
                    totalConverted++;
                }

                imageIndex++;
            }
        }

        // Validation: ensure at least one image was converted
        if (totalConverted == 0)
            throw new InvalidOperationException("No images were found and converted.");

        // Cleanup sample image file
        if (File.Exists(sampleImagePath))
            File.Delete(sampleImagePath);
    }

    // Creates a deterministic PNG image using Aspose.Drawing
    private static void CreateSamplePng(string filePath, int width, int height)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                // Draw a simple rectangle
                using (Pen pen = new Pen(Color.Blue, 5))
                {
                    g.DrawRectangle(pen, 10, 10, width - 20, height - 20);
                }
            }
            bitmap.Save(filePath, ImageFormat.Png);
        }
    }
}
