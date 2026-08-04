using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class DetectAndReplaceSmartArt
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a sample shape (not SmartArt) to have at least one shape in the document.
        // This shape will remain unchanged because it does not contain SmartArt.
        builder.InsertShape(ShapeType.Rectangle, 100, 50);

        // Traverse all shapes in the document.
        Shape[] shapes = doc.GetChildNodes(NodeType.Shape, true).OfType<Shape>().ToArray();

        foreach (Shape shape in shapes)
        {
            // Detect SmartArt shapes using the HasSmartArt property.
            if (shape.HasSmartArt)
            {
                // Create a replacement shape (simple rectangle) with the same size and position.
                Shape replacement = new Shape(doc, ShapeType.Rectangle)
                {
                    Width = shape.Width,
                    Height = shape.Height,
                    Left = shape.Left,
                    Top = shape.Top,
                    RelativeHorizontalPosition = shape.RelativeHorizontalPosition,
                    RelativeVerticalPosition = shape.RelativeVerticalPosition,
                    WrapType = shape.WrapType,
                    HorizontalAlignment = shape.HorizontalAlignment,
                    VerticalAlignment = shape.VerticalAlignment
                };

                // Insert the replacement after the original SmartArt shape and then remove the original.
                shape.ParentNode.InsertAfter(replacement, shape);
                shape.Remove();
            }
        }

        // Define the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output.docx");

        // Save the modified document.
        doc.Save(outputPath);

        // Simple validation to ensure the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved correctly.");
    }
}
