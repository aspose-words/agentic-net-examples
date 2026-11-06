using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace SmartArtReplacementExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add a paragraph to describe the content.
            builder.Writeln("Document with SmartArt placeholder:");

            // Insert a rectangle shape that will act as a SmartArt placeholder.
            // Use AlternativeText to mark it as SmartArt for detection later.
            Shape smartArtPlaceholder = builder.InsertShape(ShapeType.Rectangle, 300, 200);
            smartArtPlaceholder.AlternativeText = "SmartArt";
            smartArtPlaceholder.Stroke.Color = System.Drawing.Color.DarkBlue;
            smartArtPlaceholder.Fill.Color = System.Drawing.Color.LightBlue;

            // Save an intermediate document (optional).
            string intermediatePath = "Intermediate.docx";
            doc.Save(intermediatePath);

            // Traverse all shapes in the document.
            NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
            foreach (Shape shape in shapeNodes)
            {
                // Detect shapes that were marked as SmartArt.
                if (shape.AlternativeText == "SmartArt")
                {
                    // Create a replacement rectangle shape.
                    Shape replacement = new Shape(doc, ShapeType.Rectangle);
                    replacement.Width = shape.Width;
                    replacement.Height = shape.Height;
                    replacement.Left = shape.Left;
                    replacement.Top = shape.Top;
                    replacement.WrapType = shape.WrapType;
                    replacement.Stroke.Color = System.Drawing.Color.Black;
                    replacement.Fill.Color = System.Drawing.Color.LightGray;

                    // Insert the replacement after the original shape.
                    shape.ParentNode.InsertAfter(replacement, shape);
                    // Remove the original SmartArt placeholder.
                    shape.Remove();
                }
            }

            // Save the final document.
            string outputPath = "Output.docx";
            doc.Save(outputPath);

            // Validate that the output file was created.
            if (!File.Exists(outputPath))
                throw new Exception($"Failed to create the output file: {outputPath}");
        }
    }
}
