using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertNestedShapes
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the outermost shape (a rectangle) as an inline shape.
        Shape outerShape = builder.InsertShape(ShapeType.Rectangle, 200, 200);
        outerShape.WrapType = WrapType.Inline; // Ensure it stays inline with the text.

        // Create a middle shape (an ellipse) using the Shape constructor.
        Shape middleShape = new Shape(doc, ShapeType.Ellipse)
        {
            Width = 120,
            Height = 120,
            // Optional visual formatting.
            FillColor = System.Drawing.Color.LightBlue,
            StrokeColor = System.Drawing.Color.DarkBlue
        };

        // Append the middle shape as a child of the outer shape.
        outerShape.AppendChild(middleShape);

        // Create the innermost shape (a text box) that will hold the text.
        Shape innerShape = new Shape(doc, ShapeType.TextBox)
        {
            Width = 100,
            Height = 50,
            // Ensure the text box does not wrap text inside the shape.
            TextBox = { FitShapeToText = false }
        };

        // Append the inner shape as a child of the middle shape.
        middleShape.AppendChild(innerShape);

        // Add a paragraph to the innermost shape and insert a run of text.
        innerShape.AppendChild(new Paragraph(doc));
        Paragraph para = innerShape.FirstParagraph;
        Run run = new Run(doc)
        {
            Text = "Nested shape text"
        };
        para.AppendChild(run);

        // Save the document to a DOCX file.
        doc.Save("NestedShapes.docx");
    }
}
