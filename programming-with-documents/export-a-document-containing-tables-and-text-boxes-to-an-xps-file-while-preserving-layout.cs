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

        // Insert a simple 2‑cell table.
        builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Cell 1");
        builder.InsertCell();
        builder.Writeln("Cell 2");
        builder.EndRow();
        builder.EndTable();

        // Add some space before the text box.
        builder.Writeln();

        // Create a text box shape.
        Shape textBox = new Shape(doc, ShapeType.TextBox);
        textBox.Width = 200;
        textBox.Height = 100;
        textBox.WrapType = WrapType.Inline;

        // Add a paragraph with text inside the text box.
        Paragraph tbParagraph = new Paragraph(doc);
        Run tbRun = new Run(doc, "This is inside a text box.");
        tbParagraph.AppendChild(tbRun);
        textBox.AppendChild(tbParagraph);

        // Insert the text box into the document.
        builder.InsertNode(textBox);

        // Define output XPS file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.xps");

        // Save the document as XPS, preserving layout.
        doc.Save(outputPath, SaveFormat.Xps);

        // Verify that the XPS file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"XPS file successfully created at: {outputPath}");
        }
        else
        {
            Console.WriteLine("Failed to create the XPS file.");
        }
    }
}
