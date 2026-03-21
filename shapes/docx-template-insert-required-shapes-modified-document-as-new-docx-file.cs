using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace AsposeWordsShapeDemo
{
    class Program
    {
        static void Main()
        {
            // Use a template file if it exists; otherwise create a new blank document.
            string templatePath = Path.Combine(Environment.CurrentDirectory, "Template.docx");
            Document doc = File.Exists(templatePath) ? new Document(templatePath) : new Document();

            // Create a DocumentBuilder for editing the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // -----------------------------------------------------------------
            // Insert individual shapes.
            // -----------------------------------------------------------------
            // Insert a rectangle shape (inline) with width=150pt and height=80pt.
            Shape rect = builder.InsertShape(ShapeType.Rectangle, 150, 80);
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
            GroupShape group = builder.InsertGroupShape(rect, ellipse);
            group.Left = 50;
            group.Top = 150;

            // -----------------------------------------------------------------
            // Save the modified document to a new file in the current directory.
            // -----------------------------------------------------------------
            string outputPath = Path.Combine(Environment.CurrentDirectory, "ModifiedDocument.docx");
            doc.Save(outputPath);

            Console.WriteLine($"Document saved to: {outputPath}");
        }
    }
}
