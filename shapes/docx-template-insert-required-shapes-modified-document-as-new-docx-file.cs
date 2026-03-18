using System;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace AsposeWordsShapeDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the existing DOCX template.
            const string templatePath = @"C:\Docs\Template.docx";

            // Load the template document.
            Document doc = new Document(templatePath);

            // Create a DocumentBuilder for editing the loaded document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // -----------------------------------------------------------------
            // Insert individual shapes.
            // -----------------------------------------------------------------
            // Insert a rectangle shape (inline) with width=150pt and height=80pt.
            Shape rect = builder.InsertShape(ShapeType.Rectangle, 150, 80);
            // Optional: set fill color and line style.
            rect.FillColor = System.Drawing.Color.LightBlue;
            rect.Stroke.Color = System.Drawing.Color.DarkBlue;

            // Insert an ellipse shape (floating) positioned 100pt from the left
            // and 200pt from the top of the page, with size 120pt x 120pt.
            Shape ellipse = builder.InsertShape(
                ShapeType.Ellipse,
                RelativeHorizontalPosition.Page, 100,
                RelativeVerticalPosition.Page, 200,
                120, 120,
                WrapType.None);
            ellipse.FillColor = System.Drawing.Color.LightCoral;
            ellipse.Stroke.Color = System.Drawing.Color.Maroon;

            // -----------------------------------------------------------------
            // Group the two shapes into a single GroupShape.
            // -----------------------------------------------------------------
            // The InsertGroupShape method automatically calculates the group bounds.
            GroupShape group = builder.InsertGroupShape(rect, ellipse);

            // Optionally adjust the group's position.
            group.Left = 50;
            group.Top = 150;

            // -----------------------------------------------------------------
            // Save the modified document to a new file.
            // -----------------------------------------------------------------
            const string outputPath = @"C:\Docs\ModifiedDocument.docx";
            doc.Save(outputPath);
        }
    }
}
