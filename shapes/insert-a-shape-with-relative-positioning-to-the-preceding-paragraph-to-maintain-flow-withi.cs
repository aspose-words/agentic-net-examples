using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph with some text.
        builder.Writeln("This is the first paragraph. The shape will be positioned relative to this paragraph.");

        // Create a floating rectangle shape.
        Shape shape = new Shape(doc, ShapeType.Rectangle);
        shape.Width = 100;   // Width in points.
        shape.Height = 50;   // Height in points.
        shape.StrokeColor = Color.Blue;
        shape.FillColor = Color.LightBlue;

        // Position the shape relative to the preceding paragraph.
        shape.WrapType = WrapType.Square;
        shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Margin;
        shape.RelativeVerticalPosition = RelativeVerticalPosition.Paragraph;
        shape.Left = 0;   // 0 points from the left margin.
        shape.Top = 0;    // 0 points from the top of the preceding paragraph.

        // Append the shape to the current paragraph (the one we just wrote).
        Paragraph currentParagraph = builder.CurrentParagraph;
        currentParagraph.AppendChild(shape);

        // Save the document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");

        // Optionally, inform that the process completed.
        Console.WriteLine("Document created successfully: " + Path.GetFullPath(outputPath));
    }
}
