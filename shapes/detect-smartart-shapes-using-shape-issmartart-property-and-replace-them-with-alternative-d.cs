using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace SmartArtReplacementExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a sample shape (rectangle) – this is NOT a SmartArt shape,
            // but it allows us to demonstrate the traversal logic.
            builder.InsertShape(ShapeType.Rectangle, 100, 50);

            // Retrieve all Shape nodes in the document.
            var allShapes = doc.GetChildNodes(NodeType.Shape, true)
                               .Cast<Shape>()
                               .ToList();

            // Iterate over a copy of the list because we will modify the document tree.
            foreach (Shape shape in allShapes)
            {
                // Detect SmartArt shapes using the HasSmartArt property.
                if (shape.HasSmartArt)
                {
                    // Create a simple rectangle shape to replace the SmartArt.
                    Shape replacement = new Shape(doc, ShapeType.Rectangle)
                    {
                        Width = shape.Width,
                        Height = shape.Height,
                        Left = shape.Left,
                        Top = shape.Top,
                        RelativeHorizontalPosition = shape.RelativeHorizontalPosition,
                        RelativeVerticalPosition = shape.RelativeVerticalPosition,
                        WrapType = shape.WrapType
                    };

                    // Insert the replacement after the original SmartArt shape.
                    shape.ParentNode.InsertAfter(replacement, shape);
                    // Remove the original SmartArt shape.
                    shape.Remove();
                }
            }

            // Validation: ensure no SmartArt shapes remain in the document.
            bool anySmartArt = doc.GetChildNodes(NodeType.Shape, true)
                                  .Cast<Shape>()
                                  .Any(s => s.HasSmartArt);
            if (anySmartArt)
                throw new InvalidOperationException("Some SmartArt shapes were not replaced.");

            // Save the resulting document to the local file system.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "SmartArtReplaced.docx");
            doc.Save(outputPath);
        }
    }
}
