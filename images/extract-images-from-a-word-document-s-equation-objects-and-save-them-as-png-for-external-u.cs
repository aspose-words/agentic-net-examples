using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a deterministic sample image file.
        const string sampleImagePath = "sample.png";
        const int imgWidth = 200;
        const int imgHeight = 200;

        // Use Aspose.Drawing to create the bitmap.
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(imgWidth, imgHeight))
        {
            using (Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                // Fill background with white.
                graphics.Clear(Aspose.Drawing.Color.White);
                // Draw a simple black rectangle.
                graphics.DrawRectangle(Aspose.Drawing.Pens.Black, 20, 20, imgWidth - 40, imgHeight - 40);
            }

            // Save the bitmap to a local file.
            bitmap.Save(sampleImagePath);
        }

        // Create a Word document containing the sample image and an equation.
        const string docPath = "sample.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the sample image.
        builder.InsertImage(sampleImagePath);

        // Insert a simple equation using a field.
        builder.Writeln(); // Move to a new paragraph.
        builder.InsertField("EQ \\x\\to(\\int)", null);

        // Save the document.
        doc.Save(docPath);

        // Reload the document for extraction.
        Document loadedDoc = new Document(docPath);

        // Extract images from all shape nodes (including those that may belong to equations).
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes)
        {
            if (shape.HasImage)
            {
                string outputImagePath = $"equation-image-{imageIndex}.png";

                // Save the image data to a PNG file.
                shape.ImageData.Save(outputImagePath);

                // Validate that the file was created.
                if (!File.Exists(outputImagePath))
                {
                    throw new InvalidOperationException($"Failed to save extracted image to '{outputImagePath}'.");
                }

                imageIndex++;
            }
        }

        // Validate that at least one image was extracted.
        if (imageIndex == 0)
        {
            throw new InvalidOperationException("No images were extracted from the document.");
        }

        // Optional cleanup (commented out).
        // File.Delete(sampleImagePath);
        // File.Delete(docPath);
    }
}
