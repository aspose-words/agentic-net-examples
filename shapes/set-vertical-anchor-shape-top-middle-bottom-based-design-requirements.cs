using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a textbox shape with a specific size.
        Shape shape = builder.InsertShape(ShapeType.TextBox, 200, 100);

        // Set the vertical anchor of the textbox (Top, Middle, or Bottom).
        SetVerticalAnchor(shape, TextBoxAnchor.Middle);

        // Add some text inside the textbox.
        builder.MoveTo(shape.FirstParagraph);
        builder.Write("Hello world!");

        // Compatibility option for Word 2007 is optional for vertical anchoring.
        // If you need it and your Aspose.Words version supports MsWordVersion,
        // uncomment the line below and ensure the proper using directive.
        // doc.CompatibilityOptions.OptimizeFor(MsWordVersion.Word2007);

        // Save the document to disk.
        doc.Save("Shape.VerticalAnchor.docx");
    }

    // Helper method to assign the vertical anchor to a textbox shape.
    static void SetVerticalAnchor(Shape shape, TextBoxAnchor anchor)
    {
        // The TextBox property provides access to textbox-specific formatting.
        shape.TextBox.VerticalAnchor = anchor;
    }
}
