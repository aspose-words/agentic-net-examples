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
        // Create a deterministic sample image that will act as audio cover art.
        const string coverImagePath = "cover.png";
        CreateSampleImage(coverImagePath);

        // Build a DOCX document and insert the sample image.
        const string docPath = "sample.docx";
        CreateDocumentWithImage(docPath, coverImagePath);

        // Load the document and extract all images, saving them as JPEG files.
        ExtractImagesAsJpeg(docPath);
    }

    private static void CreateSampleImage(string filePath)
    {
        // Create a 200x200 white bitmap.
        using (Bitmap bitmap = new Bitmap(200, 200))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(Color.White);
            // Additional deterministic drawing can be added here if needed.
            bitmap.Save(filePath);
        }
    }

    private static void CreateDocumentWithImage(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the image into the document (simulating audio cover art).
        builder.InsertImage(imagePath);

        // Save the document.
        doc.Save(docPath);
    }

    private static void ExtractImagesAsJpeg(string docPath)
    {
        Document doc = new Document(docPath);

        // Get all shape nodes in the document.
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        int imageIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                // Save each image as a JPEG file with a deterministic name.
                string outputFileName = $"extracted_{imageIndex}.jpg";
                shape.ImageData.Save(outputFileName);
                imageIndex++;
            }
        }

        // Validate that at least one image was extracted.
        if (imageIndex == 0)
            throw new InvalidOperationException("No images were extracted from the document.");
    }
}
