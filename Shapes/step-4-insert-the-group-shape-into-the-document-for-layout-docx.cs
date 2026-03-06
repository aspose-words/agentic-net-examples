using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using System.Drawing;

namespace AsposeWordsGroupShapeExample
{
    class Program
    {
        static void Main()
        {
            // Step 1. Create a new blank document.
            Document doc = new Document();

            // Step 2. Initialize a DocumentBuilder for inserting content.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Step 3. Insert two individual shapes that will later be grouped.
            // Rectangle shape.
            Shape rect = builder.InsertShape(ShapeType.Rectangle, 200, 250);
            rect.Left = 20;   // Position relative to the page.
            rect.Top = 20;
            rect.Stroke.Color = Color.Red;

            // Ellipse shape.
            Shape ellipse = builder.InsertShape(ShapeType.Ellipse, 150, 200);
            ellipse.Left = 40;
            ellipse.Top = 50;
            ellipse.Stroke.Color = Color.Green;

            // Step 4. Group the previously inserted shapes.
            // The InsertGroupShape method automatically calculates the group’s position and size.
            GroupShape group = builder.InsertGroupShape(rect, ellipse);

            // Optional: add a caption inside the group to demonstrate text handling.
            Paragraph para = new Paragraph(doc);
            Run run = new Run(doc, "Grouped Shapes");
            para.AppendChild(run);
            // Append the paragraph to the first child shape inside the group.
            ((Shape)group.GetChild(NodeType.Shape, 0, true)).AppendChild(para);

            // Step 5. Save the document in DOCX format.
            doc.Save("GroupedShapes.docx");
        }
    }
}
