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
        // Deterministic file names
        const string previewImagePath = "preview.png";
        const string docPath = "sample.docx";

        // -------------------------------------------------
        // Step 1: Create a sample preview image (200x200)
        // -------------------------------------------------
        using (Bitmap bitmap = new Bitmap(200, 200))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                using (Pen pen = new Pen(Color.Blue, 5))
                {
                    g.DrawRectangle(pen, 10, 10, 180, 180);
                }
            }
            bitmap.Save(previewImagePath);
        }

        // -------------------------------------------------
        // Step 2: Build a DOCX containing an image that
        // simulates a video preview frame
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create an image shape (used as a video preview placeholder)
        Shape imageShape = new Shape(doc, ShapeType.Image);
        imageShape.Width = 300.0;
        imageShape.Height = 200.0;

        // Assign the preview image to the shape
        imageShape.ImageData.SetImage(previewImagePath);

        // Append the shape to the document
        builder.CurrentParagraph.AppendChild(imageShape);

        // Save the document
        doc.Save(docPath);

        // -------------------------------------------------
        // Step 3: Load the document and extract preview images
        // -------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        int extractedCount = 0;
        int index = 0;
        foreach (Shape shape in shapeNodes)
        {
            // Look for image shapes that have image data
            if (shape.HasImage)
            {
                string outputImagePath = $"extracted-{index}.png";
                shape.ImageData.Save(outputImagePath);
                if (!File.Exists(outputImagePath))
                {
                    throw new InvalidOperationException($"Failed to save extracted image to {outputImagePath}");
                }
                extractedCount++;
                index++;
            }
        }

        // Validation: ensure at least one image was extracted
        if (extractedCount == 0)
        {
            throw new InvalidOperationException("No preview images were extracted from the document.");
        }
    }
}
