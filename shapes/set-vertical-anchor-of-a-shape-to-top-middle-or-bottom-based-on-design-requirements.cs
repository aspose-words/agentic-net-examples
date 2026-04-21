using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class SetShapeVerticalAnchorExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Define common shape size.
        double shapeWidth = 200;
        double shapeHeight = 100;

        // Insert first text box with vertical anchor at the top.
        Shape topBox = builder.InsertShape(ShapeType.TextBox, shapeWidth, shapeHeight);
        topBox.WrapType = WrapType.None;
        topBox.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        topBox.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        topBox.Left = 50;
        topBox.Top = 50;
        topBox.TextBox.VerticalAnchor = TextBoxAnchor.Top;
        builder.MoveTo(topBox.FirstParagraph);
        builder.Write("Top anchor");

        // Insert second text box with vertical anchor in the middle.
        Shape middleBox = builder.InsertShape(ShapeType.TextBox, shapeWidth, shapeHeight);
        middleBox.WrapType = WrapType.None;
        middleBox.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        middleBox.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        middleBox.Left = 300;
        middleBox.Top = 50;
        middleBox.TextBox.VerticalAnchor = TextBoxAnchor.Middle;
        builder.MoveTo(middleBox.FirstParagraph);
        builder.Write("Middle anchor");

        // Insert third text box with vertical anchor at the bottom.
        Shape bottomBox = builder.InsertShape(ShapeType.TextBox, shapeWidth, shapeHeight);
        bottomBox.WrapType = WrapType.None;
        bottomBox.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        bottomBox.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        bottomBox.Left = 550;
        bottomBox.Top = 50;
        bottomBox.TextBox.VerticalAnchor = TextBoxAnchor.Bottom;
        builder.MoveTo(bottomBox.FirstParagraph);
        builder.Write("Bottom anchor");

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ShapeVerticalAnchor.docx");
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved successfully.");
    }
}
