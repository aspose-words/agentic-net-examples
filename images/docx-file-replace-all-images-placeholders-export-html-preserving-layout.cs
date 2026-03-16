using System;
using System.IO;
using System.Linq;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the DOCX file from disk.
        Document doc = new Document("input.docx");

        // Retrieve all Shape nodes (including images) in the document.
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        // Iterate over the shapes. Use a copy of the collection because we will modify the tree.
        foreach (Shape shape in shapeNodes.OfType<Shape>().ToList())
        {
            // Process only shapes that contain an image.
            if (shape.HasImage)
            {
                // Create a rectangle shape that will serve as a placeholder.
                Shape placeholder = new Shape(doc, ShapeType.Rectangle);

                // Preserve the original image's size and positioning.
                placeholder.Width = shape.Width;
                placeholder.Height = shape.Height;
                placeholder.WrapType = shape.WrapType;
                placeholder.RelativeHorizontalPosition = shape.RelativeHorizontalPosition;
                placeholder.RelativeVerticalPosition = shape.RelativeVerticalPosition;
                placeholder.HorizontalAlignment = shape.HorizontalAlignment;
                placeholder.VerticalAlignment = shape.VerticalAlignment;
                placeholder.Left = shape.Left;
                placeholder.Top = shape.Top;

                // Visual styling for the placeholder.
                placeholder.FillColor = Color.LightGray;
                placeholder.StrokeColor = Color.DarkGray;

                // Add descriptive text inside the placeholder.
                // Use a Paragraph + Run because Shape.TextPath is read‑only for Rectangle shapes.
                Paragraph para = new Paragraph(doc);
                Run run = new Run(doc, "Image Placeholder");
                run.Font.Name = "Arial";
                run.Font.Size = 12;
                run.Font.Bold = true;
                para.AppendChild(run);
                placeholder.AppendChild(para);

                // Insert the placeholder after the original image shape and then remove the image.
                shape.ParentNode.InsertAfter(placeholder, shape);
                shape.Remove();
            }
        }

        // Save the modified document as HTML, preserving the original layout.
        doc.Save("output.html", SaveFormat.Html);
    }
}
