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

        // Insert the preceding paragraph.
        builder.Writeln("This is the first paragraph. The shape will be positioned relative to this paragraph.");

        // Insert a floating rectangle shape anchored to the preceding paragraph.
        Shape shape = builder.InsertShape(ShapeType.Rectangle, 100, 50);
        shape.WrapType = WrapType.Square;
        shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Margin;
        shape.RelativeVerticalPosition = RelativeVerticalPosition.Paragraph;
        shape.Left = 0; // Position relative to the left margin of the paragraph.
        shape.Top = 0;  // Position at the top of the paragraph.
        shape.StrokeColor = Color.Blue;
        shape.FillColor = Color.LightBlue;

        // Save the document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");
    }
}
