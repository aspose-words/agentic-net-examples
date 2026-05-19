using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class SetVerticalAnchorExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a textbox shape and set its vertical anchor to Top.
        Shape topBox = builder.InsertShape(ShapeType.TextBox, 200, 100);
        topBox.TextBox.VerticalAnchor = TextBoxAnchor.Top;
        builder.MoveTo(topBox.FirstParagraph);
        builder.Write("Top anchor");

        // Insert a second textbox shape and set its vertical anchor to Middle.
        Shape middleBox = builder.InsertShape(ShapeType.TextBox, 200, 100);
        middleBox.TextBox.VerticalAnchor = TextBoxAnchor.Middle;
        builder.MoveTo(middleBox.FirstParagraph);
        builder.Write("Middle anchor");

        // Insert a third textbox shape and set its vertical anchor to Bottom.
        Shape bottomBox = builder.InsertShape(ShapeType.TextBox, 200, 100);
        bottomBox.TextBox.VerticalAnchor = TextBoxAnchor.Bottom;
        builder.MoveTo(bottomBox.FirstParagraph);
        builder.Write("Bottom anchor");

        // Save the document to the current directory.
        string fileName = Path.Combine(Directory.GetCurrentDirectory(), "VerticalAnchor.docx");
        doc.Save(fileName);

        // Validate that the file was created.
        if (!File.Exists(fileName))
            throw new InvalidOperationException("The output document was not saved correctly.");
    }
}
