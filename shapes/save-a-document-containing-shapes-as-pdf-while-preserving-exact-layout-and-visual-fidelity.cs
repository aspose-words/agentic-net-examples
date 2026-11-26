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

        // Insert a paragraph to hold the shapes.
        builder.Writeln("Document with shapes saved as PDF:");
        builder.Writeln();

        // Insert a rectangle shape (floating) with explicit size and position.
        Shape rectangle = new Shape(doc, ShapeType.Rectangle);
        rectangle.Width = 200;
        rectangle.Height = 100;
        rectangle.Left = 100;   // points from the left margin
        rectangle.Top = 50;     // points from the top margin
        rectangle.WrapType = WrapType.None; // floating shape
        rectangle.StrokeColor = System.Drawing.Color.Blue;
        rectangle.FillColor = System.Drawing.Color.LightYellow;

        // Append the rectangle to the document body.
        doc.FirstSection.Body.FirstParagraph.AppendChild(rectangle);

        // Insert a text box shape (floating) with some text inside.
        Shape textBox = new Shape(doc, ShapeType.TextBox);
        textBox.Width = 250;
        textBox.Height = 80;
        textBox.Left = 350;
        textBox.Top = 150;
        textBox.WrapType = WrapType.Square;
        textBox.StrokeColor = System.Drawing.Color.Green;
        textBox.FillColor = System.Drawing.Color.LightCyan;

        // Add text to the text box.
        Paragraph tbParagraph = new Paragraph(doc);
        Run tbRun = new Run(doc, "This is a text box shape.");
        tbParagraph.AppendChild(tbRun);
        textBox.AppendChild(tbParagraph);

        // Append the text box to the document.
        doc.FirstSection.Body.FirstParagraph.AppendChild(textBox);

        // Save the document as PDF.
        string outputPath = "ShapesOutput.pdf";
        doc.Save(outputPath, SaveFormat.Pdf);

        // Validate that the PDF file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception($"Failed to create PDF file at '{outputPath}'.");
        }

        // Optionally, clean up (not required for this example).
        // Console.WriteLine("PDF saved successfully.");
    }
}
