using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a simple rectangle shape as a placeholder.
        // (SmartArt insertion is omitted because the required namespace is unavailable in this version.)
        Shape placeholder = builder.InsertShape(ShapeType.Rectangle, 400, 300);
        placeholder.WrapType = WrapType.None;

        // Traverse all shapes in the document.
        var shapes = doc.GetChildNodes(NodeType.Shape, true).Cast<Shape>().ToList();

        foreach (var shape in shapes)
        {
            // Detect SmartArt shapes using the HasSmartArt property.
            if (shape.HasSmartArt)
            {
                // Create a replacement rectangle shape with the same size and position.
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

                // Insert the replacement after the original SmartArt shape and remove the original.
                shape.ParentNode.InsertAfter(replacement, shape);
                shape.Remove();
            }
        }

        // Define the output path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "SmartArt_Replaced.docx");

        // Save the modified document.
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved correctly.");
    }
}
