using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a text box with vertical anchor at the top.
        Shape topBox = builder.InsertShape(ShapeType.TextBox, 200, 100);
        topBox.TextBox.VerticalAnchor = TextBoxAnchor.Top;
        builder.MoveTo(topBox.FirstParagraph);
        builder.Write("Top anchor");

        // Insert a text box with vertical anchor in the middle.
        Shape middleBox = builder.InsertShape(ShapeType.TextBox, 200, 100);
        middleBox.TextBox.VerticalAnchor = TextBoxAnchor.Middle;
        builder.MoveTo(middleBox.FirstParagraph);
        builder.Write("Middle anchor");

        // Insert a text box with vertical anchor at the bottom.
        Shape bottomBox = builder.InsertShape(ShapeType.TextBox, 200, 100);
        bottomBox.TextBox.VerticalAnchor = TextBoxAnchor.Bottom;
        builder.MoveTo(bottomBox.FirstParagraph);
        builder.Write("Bottom anchor");

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "VerticalAnchor.docx");
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not saved successfully.");
    }
}
