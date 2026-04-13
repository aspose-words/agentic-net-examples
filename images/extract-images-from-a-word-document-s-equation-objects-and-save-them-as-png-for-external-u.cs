using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a sample PNG image.
        const string inputImagePath = "input.png";
        using (Bitmap bitmap = new Bitmap(100, 100))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.White);
                using (Pen pen = new Pen(Color.Blue, 3))
                {
                    graphics.DrawRectangle(pen, 10, 10, 80, 80);
                }
            }
            bitmap.Save(inputImagePath);
        }

        // Verify the sample image was created.
        if (!File.Exists(inputImagePath))
            throw new Exception($"Failed to create sample image at '{inputImagePath}'.");

        // Step 2: Create a Word document, insert the image as a shape and add an equation.
        const string docPath = "sample.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the image as a shape.
        Shape imageShape = new Shape(doc, ShapeType.Image);
        imageShape.ImageData.SetImage(inputImagePath);
        builder.InsertNode(imageShape);
        builder.Writeln();

        // Insert a simple equation.
        builder.InsertField(@"EQ \o(\a,\b)");
        builder.Writeln();

        // Save the document.
        doc.Save(docPath);

        // Verify the document was saved.
        if (!File.Exists(docPath))
            throw new Exception($"Failed to create Word document at '{docPath}'.");

        // Step 3: Load the document and extract images from shape nodes.
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        int extractedCount = 0;
        foreach (Shape shape in shapeNodes)
        {
            if (shape.HasImage)
            {
                extractedCount++;
                string outputImagePath = $"extracted-{extractedCount}.png";
                shape.ImageData.Save(outputImagePath);

                // Validate that the image was saved.
                if (!File.Exists(outputImagePath))
                    throw new Exception($"Failed to save extracted image to '{outputImagePath}'.");
            }
        }

        // Ensure at least one image was extracted.
        if (extractedCount == 0)
            throw new Exception("No images were extracted from the document.");

        // Optional: Clean up created files (comment out if you want to keep them).
        // File.Delete(inputImagePath);
        // File.Delete(docPath);
        // for (int i = 1; i <= extractedCount; i++)
        //     File.Delete($"extracted-{i}.png");
    }
}
