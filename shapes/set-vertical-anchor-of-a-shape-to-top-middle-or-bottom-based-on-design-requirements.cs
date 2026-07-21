using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class SetShapeVerticalAnchor
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Define common size for the text boxes.
        const double boxWidth = 200;
        const double boxHeight = 100;

        // Insert first text box with vertical anchor at the top.
        Shape topBox = builder.InsertShape(ShapeType.TextBox, boxWidth, boxHeight);
        topBox.TextBox.VerticalAnchor = TextBoxAnchor.Top;
        builder.MoveTo(topBox.FirstParagraph);
        builder.Write("Top anchor");
        // Move cursor below the shape to continue building.
        builder.Writeln();

        // Insert second text box with vertical anchor in the middle.
        Shape middleBox = builder.InsertShape(ShapeType.TextBox, boxWidth, boxHeight);
        middleBox.TextBox.VerticalAnchor = TextBoxAnchor.Middle;
        builder.MoveTo(middleBox.FirstParagraph);
        builder.Write("Middle anchor");
        builder.Writeln();

        // Insert third text box with vertical anchor at the bottom.
        Shape bottomBox = builder.InsertShape(ShapeType.TextBox, boxWidth, boxHeight);
        bottomBox.TextBox.VerticalAnchor = TextBoxAnchor.Bottom;
        builder.MoveTo(bottomBox.FirstParagraph);
        builder.Write("Bottom anchor");
        builder.Writeln();

        // Save the document to the local file system.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "VerticalAnchor.docx");
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");

        // Optionally, inform that the process completed successfully.
        Console.WriteLine($"Document saved successfully to: {outputPath}");
    }
}
