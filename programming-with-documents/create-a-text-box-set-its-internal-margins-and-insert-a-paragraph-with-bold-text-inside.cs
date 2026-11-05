using System;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a text box shape.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 100);

        // Set internal margins (in points).
        textBox.TextBox.InternalMarginTop = 10;
        textBox.TextBox.InternalMarginBottom = 10;
        textBox.TextBox.InternalMarginLeft = 10;
        textBox.TextBox.InternalMarginRight = 10;

        // Move the builder inside the text box and add a bold paragraph.
        builder.MoveTo(textBox.FirstParagraph);
        builder.Font.Bold = true;
        builder.Writeln("Bold text inside the text box.");

        // Save the document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Indicate completion.
        Console.WriteLine($"Document saved to {outputPath}");
    }
}
