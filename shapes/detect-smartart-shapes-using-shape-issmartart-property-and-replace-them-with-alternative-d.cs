using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with a placeholder shape (non‑SmartArt).
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        Shape placeholderShape = builder.InsertShape(ShapeType.Rectangle, 150, 100);
        placeholderShape.WrapType = WrapType.Inline;
        placeholderShape.Stroke.Color = System.Drawing.Color.Blue;
        placeholderShape.FillColor = System.Drawing.Color.LightGray;

        // Save the sample document (input).
        const string inputPath = "Input.docx";
        doc.Save(inputPath);

        // Load the document (could be any document that may contain SmartArt).
        Document loadedDoc = new Document(inputPath);

        // Prepare a replacement image (generated programmatically).
        const string replacementImagePath = "ReplacementDiagram.png";
        GeneratePlaceholderImage(replacementImagePath, 150, 100);

        // Traverse all shapes in the document.
        var shapes = loadedDoc.GetChildNodes(NodeType.Shape, true).OfType<Shape>().ToList();

        foreach (Shape shape in shapes)
        {
            // Detect SmartArt shapes using the HasSmartArt property.
            if (shape.HasSmartArt)
            {
                // Create a new image shape to replace the SmartArt.
                Shape imageShape = new Shape(loadedDoc, ShapeType.Image);
                imageShape.ImageData.SetImage(replacementImagePath);

                // Preserve the original shape's size and position.
                imageShape.Width = shape.Width;
                imageShape.Height = shape.Height;
                imageShape.Left = shape.Left;
                imageShape.Top = shape.Top;
                imageShape.RelativeHorizontalPosition = shape.RelativeHorizontalPosition;
                imageShape.RelativeVerticalPosition = shape.RelativeVerticalPosition;
                imageShape.WrapType = shape.WrapType;
                imageShape.WrapSide = shape.WrapSide;
                imageShape.HorizontalAlignment = shape.HorizontalAlignment;
                imageShape.VerticalAlignment = shape.VerticalAlignment;

                // Insert the new image shape after the original SmartArt shape.
                shape.ParentNode.InsertAfter(imageShape, shape);
                // Remove the original SmartArt shape.
                shape.Remove();
            }
        }

        // Save the modified document.
        const string outputPath = "Output.docx";
        loadedDoc.Save(outputPath);

        // Simple validation to ensure the output file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved correctly.");

        // Clean up temporary files.
        if (File.Exists(inputPath)) File.Delete(inputPath);
        if (File.Exists(replacementImagePath)) File.Delete(replacementImagePath);
    }

    // Generates a minimal PNG file (1x1 pixel) as a placeholder image.
    private static void GeneratePlaceholderImage(string filePath, int width, int height)
    {
        // This is a base64‑encoded 1×1 pixel transparent PNG.
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X9WcAAAAASUVORK5CYII=";
        byte[] pngBytes = Convert.FromBase64String(base64Png);
        File.WriteAllBytes(filePath, pngBytes);
    }
}
