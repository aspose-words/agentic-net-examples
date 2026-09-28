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

        // Create a text box shape.
        Shape textBox = new Shape(doc, ShapeType.TextBox);
        textBox.Width = 200;   // Width in points.
        textBox.Height = 100;  // Height in points.

        // Set the anchor position to be relative to the page margin.
        textBox.RelativeHorizontalPosition = RelativeHorizontalPosition.Margin;
        textBox.RelativeVerticalPosition = RelativeVerticalPosition.Margin;
        textBox.Left = 0;   // Distance from the left margin.
        textBox.Top = 0;    // Distance from the top margin.

        // Add a paragraph with some text inside the text box.
        Paragraph para = new Paragraph(doc);
        Run run = new Run(doc, "Hello from the text box!");
        para.AppendChild(run);
        // For a TextBox shape, content is added directly to the shape's child nodes.
        textBox.AppendChild(para);

        // Insert the text box into the document.
        builder.InsertNode(textBox);

        // Save the document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"Document saved successfully to '{outputPath}'.");
        }
        else
        {
            Console.WriteLine("Failed to save the document.");
        }
    }
}
