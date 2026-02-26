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
        Shape outerShape = builder.InsertShape(ShapeType.Rectangle, 300, 200);
        outerShape.WrapType = WrapType.Inline;

        // Create the middle shape (an ellipse) and set its size.
        Shape middleShape = new Shape(doc, ShapeType.Ellipse)
        {
            Width = 150,
            Height = 100,
            WrapType = WrapType.Inline
        };

        // Append the middle shape as a child of the outer shape.
        outerShape.AppendChild(middleShape);

        // Create the innermost shape (a text box) and set its size.
        Shape innerShape = new Shape(doc, ShapeType.TextBox)
        {
            Width = 120,
            Height = 60,
            WrapType = WrapType.Inline
        };

        // Append the inner shape as a child of the middle shape.
        middleShape.AppendChild(innerShape);

        // Add a paragraph and a run of text inside the innermost shape.
        Paragraph para = innerShape.FirstParagraph;
        Run run = new Run(doc) { Text = "Nested shape text" };
        para.AppendChild(run);

        // Save the document to a DOCX file.
        doc.Save("NestedShapes.docx");
    }
}
