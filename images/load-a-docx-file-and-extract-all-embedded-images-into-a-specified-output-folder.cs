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
        // Define base directories.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        string imagesDir = Path.Combine(baseDir, "ExtractedImages");
        Directory.CreateDirectory(baseDir);
        Directory.CreateDirectory(imagesDir);

        // Create a deterministic sample image (sample.png).
        string sampleImagePath = Path.Combine(baseDir, "sample.png");
        CreateSampleImage(sampleImagePath, 200, 200);

        // Create a DOCX document and insert the sample image.
        string docPath = Path.Combine(baseDir, "sample.docx");
        CreateDocumentWithImage(docPath, sampleImagePath);

        // Load the document and extract all embedded images.
        ExtractImagesFromDocument(docPath, imagesDir);
    }

    private static void CreateSampleImage(string filePath, int width, int height)
    {
        // Create a bitmap and clear it with white color.
        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(Color.White);
            // Optionally, draw something deterministic (a black rectangle).
            graphics.DrawRectangle(new Pen(Color.Black, 2), 10, 10, width - 20, height - 20);
            // Save the bitmap to the specified file.
            bitmap.Save(filePath);
        }
    }

    private static void CreateDocumentWithImage(string docPath, string imagePath)
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the image into the document.
        builder.InsertImage(imagePath);

        // Save the document.
        doc.Save(docPath);
    }

    private static void ExtractImagesFromDocument(string docPath, string outputFolder)
    {
        // Load the document.
        Document doc = new Document(docPath);

        // Get all shape nodes (including images).
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

        int imageIndex = 0;
        foreach (Shape shape in shapes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                // Determine the appropriate file extension for the image type.
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string outputPath = Path.Combine(outputFolder, $"extracted_{imageIndex}{extension}");

                // Save the image data to the file system.
                shape.ImageData.Save(outputPath);
                imageIndex++;
            }
        }

        // Validate that at least one image was extracted.
        if (imageIndex == 0)
        {
            throw new Exception("No images were extracted from the document.");
        }

        Console.WriteLine($"Extracted {imageIndex} image(s) to \"{outputFolder}\".");
    }
}
