using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class SetShapeAlternativeText
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an inline cube shape with a size of 150x150 points.
        Shape shape = builder.InsertShape(ShapeType.Cube, 150, 150);

        // Give the shape a name (optional, helps identify the shape in the document).
        shape.Name = "MyCube";

        // Set the alternative text – this text is read by screen readers and shown when the image is missing.
        shape.AlternativeText = "A 3‑D cube used as a decorative illustration.";

        // Save the document to a DOCX file.
        doc.Save("ShapeWithAltText.docx");
    }
}
