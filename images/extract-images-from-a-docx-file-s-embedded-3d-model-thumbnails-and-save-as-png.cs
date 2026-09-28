using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a deterministic sample image to be used as a thumbnail.
        const string inputImagePath = "input.png";
        const int width = 200;
        const int height = 200;

        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                // Fill background with white.
                graphics.Clear(Color.White);
                // Draw a simple red rectangle.
                graphics.FillRectangle(new SolidBrush(Color.Red), 20, 20, width - 40, height - 40);
            }

            // Save the sample image.
            bitmap.Save(inputImagePath);
        }

        // Verify that the sample image was created.
        if (!File.Exists(inputImagePath))
            throw new Exception($"Failed to create sample image '{inputImagePath}'.");

        // Create a new document and insert the sample image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(inputImagePath);
        const string docPath = "sample.docx";
        doc.Save(docPath);

        // Verify that the document was saved.
        if (!File.Exists(docPath))
            throw new Exception($"Failed to save document '{docPath}'.");

        // Load the document for extraction.
        Document loadedDoc = new Document(docPath);

        // Iterate through all Shape nodes and extract images (thumbnails).
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;

        foreach (Shape shape in shapeNodes)
        {
            if (shape.HasImage)
            {
                string outputImagePath = $"thumbnail-{extractedCount}.png";
                shape.ImageData.Save(outputImagePath);

                // Validate that the image file was created.
                if (!File.Exists(outputImagePath))
                    throw new Exception($"Failed to save extracted image '{outputImagePath}'.");

                extractedCount++;
            }
        }

        // Ensure at least one image was extracted.
        if (extractedCount == 0)
            throw new Exception("No images were extracted from the document.");

        // Cleanup: optional removal of temporary files (commented out to keep outputs).
        // File.Delete(inputImagePath);
        // File.Delete(docPath);
    }
}
