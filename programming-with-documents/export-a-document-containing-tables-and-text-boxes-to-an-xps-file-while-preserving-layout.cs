using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a 2x2 table.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Cell 1");
        builder.InsertCell();
        builder.Writeln("Cell 2");
        builder.EndRow();

        builder.InsertCell();
        builder.Writeln("Cell 3");
        builder.InsertCell();
        builder.Writeln("Cell 4");
        builder.EndRow();
        builder.EndTable();

        // Insert a text box (shape with a TextBox).
        Shape textBox = new Shape(doc, ShapeType.TextBox);
        textBox.Width = 200;
        textBox.Height = 100;
        textBox.WrapType = WrapType.Inline;

        // Add text to the text box.
        Paragraph tbParagraph = new Paragraph(doc);
        Run tbRun = new Run(doc, "This is a text box.");
        tbParagraph.AppendChild(tbRun);
        textBox.AppendChild(tbParagraph);

        // Place the text box after the table.
        builder.InsertNode(textBox);

        // Save the document as XPS.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output.xps");
        doc.Save(outputPath, SaveFormat.Xps);

        // Verify that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine("XPS file created successfully at: " + outputPath);
        }
        else
        {
            Console.WriteLine("Failed to create XPS file.");
        }
    }
}
