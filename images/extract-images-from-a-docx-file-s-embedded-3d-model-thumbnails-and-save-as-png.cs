using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Extract3DModelThumbnails
{
    public static void Main()
    {
        // Prepare a deterministic sample image to act as a 3D model thumbnail.
        const string thumbnailPath = "thumb.png";
        const int width = 200;
        const int height = 200;

        // Create the sample thumbnail image.
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.White);
                // Draw a simple red ellipse to make the thumbnail recognizable.
                graphics.FillEllipse(new SolidBrush(Color.Red), 20, 20, width - 40, height - 40);
            }
            bitmap.Save(thumbnailPath);
        }

        // Create a new Word document and insert the thumbnail as a shape.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a shape that holds the image (simulating a 3D model thumbnail).
        Shape thumbnailShape = new Shape(doc, ShapeType.Image);
        thumbnailShape.ImageData.SetImage(thumbnailPath);
        // Append the shape to the current paragraph.
        builder.CurrentParagraph.AppendChild(thumbnailShape);

        // Save the document containing the embedded thumbnail.
        const string docPath = "sample.docx";
        doc.Save(docPath);

        // Load the document for extraction.
        Document loadedDoc = new Document(docPath);

        // Iterate through all Shape nodes and extract images.
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;
        int index = 0;

        foreach (Shape shape in shapeNodes)
        {
            if (shape.HasImage)
            {
                string outputImagePath = $"extracted-{index}.png";
                shape.ImageData.Save(outputImagePath);
                extractedCount++;
                index++;
            }
        }

        // Validate that at least one image was extracted.
        if (extractedCount == 0)
        {
            throw new InvalidOperationException("No images were extracted from the document.");
        }

        // Optional: Clean up temporary files (comment out if you want to keep them).
        // File.Delete(thumbnailPath);
        // File.Delete(docPath);
    }
}
