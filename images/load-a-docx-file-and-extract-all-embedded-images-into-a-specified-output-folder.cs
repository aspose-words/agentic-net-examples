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
        // Base directory for all temporary files.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        Directory.CreateDirectory(baseDir);

        // -----------------------------------------------------------------
        // 1. Create a deterministic sample image (input.png).
        // -----------------------------------------------------------------
        string sampleImagePath = Path.Combine(baseDir, "input.png");
        CreateSampleImage(sampleImagePath, 200, 200);

        // -----------------------------------------------------------------
        // 2. Create a sample DOCX that contains the image.
        // -----------------------------------------------------------------
        string docPath = Path.Combine(baseDir, "sample.docx");
        CreateSampleDocument(docPath, sampleImagePath);

        // -----------------------------------------------------------------
        // 3. Load the DOCX file.
        // -----------------------------------------------------------------
        Document doc = new Document(docPath);

        // -----------------------------------------------------------------
        // 4. Extract all embedded images to the output folder.
        // -----------------------------------------------------------------
        string outputFolder = Path.Combine(baseDir, "ExtractedImages");
        Directory.CreateDirectory(outputFolder);

        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                // Determine the proper file extension for the image type.
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string outputPath = Path.Combine(outputFolder, $"image_{imageIndex}{extension}");

                // Save the image data to the file system.
                shape.ImageData.Save(outputPath);
                imageIndex++;
            }
        }

        // Validate that at least one image was extracted.
        if (imageIndex == 0)
            throw new Exception("No images were extracted from the document.");

        // Optional: indicate success.
        Console.WriteLine($"Extracted {imageIndex} image(s) to \"{outputFolder}\".");
    }

    // Creates a simple white bitmap with optional drawing.
    private static void CreateSampleImage(string filePath, int width, int height)
    {
        // Create bitmap.
        Bitmap bitmap = new Bitmap(width, height);
        // Obtain graphics object.
        Graphics graphics = Graphics.FromImage(bitmap);
        // Fill with white background.
        graphics.Clear(Color.White);
        // (Additional deterministic drawing can be added here if desired.)

        // Save bitmap to file.
        bitmap.Save(filePath);

        // Clean up resources.
        graphics.Dispose();
        bitmap.Dispose();
    }

    // Creates a DOCX file that contains the specified image.
    private static void CreateSampleDocument(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the image into the document.
        builder.InsertImage(imagePath);

        // Save the document.
        doc.Save(docPath);
    }
}
