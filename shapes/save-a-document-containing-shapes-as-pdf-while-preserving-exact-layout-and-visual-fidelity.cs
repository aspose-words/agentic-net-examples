using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a rectangle shape.
        Shape rectangle = builder.InsertShape(ShapeType.Rectangle, 150, 100);
        rectangle.StrokeColor = Color.Blue;
        rectangle.FillColor = Color.LightBlue;
        rectangle.WrapType = WrapType.None; // Keep exact layout.

        // Insert a line break.
        builder.Writeln();

        // Insert an ellipse shape.
        Shape ellipse = builder.InsertShape(ShapeType.Ellipse, 120, 80);
        ellipse.StrokeColor = Color.Green;
        ellipse.FillColor = Color.LightGreen;
        ellipse.WrapType = WrapType.None;

        // Insert a line break.
        builder.Writeln();

        // Insert a text box shape.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 100);
        textBox.WrapType = WrapType.None;
        // Add text inside the text box.
        builder.MoveTo(textBox.FirstParagraph);
        builder.Writeln("Sample text inside a text box.");

        // Define output PDF path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ShapesOutput.pdf");

        // Save the document as PDF.
        doc.Save(outputPath, SaveFormat.Pdf);

        // Validate that the PDF file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception("Failed to create the PDF file.");
        }

        // Optionally, clean up (not required for the task).
        // Console.WriteLine($"PDF saved to: {outputPath}");
    }
}
