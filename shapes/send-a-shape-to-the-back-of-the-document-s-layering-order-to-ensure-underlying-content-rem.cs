using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph of text that will be overlapped by a shape.
        builder.Writeln("This text will be overlapped by a shape, but the shape will be sent to the back.");

        // Insert a floating rectangle shape.
        Shape shape = builder.InsertShape(ShapeType.Rectangle, 200, 100);
        shape.WrapType = WrapType.None;
        shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        shape.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        shape.Left = 0;
        shape.Top = 0;
        shape.FillColor = System.Drawing.Color.LightBlue;
        shape.StrokeColor = System.Drawing.Color.DarkBlue;

        // Send the shape to the back of the layering order (behind the text).
        shape.BehindText = true;

        // Save the document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not saved correctly.");
    }
}
