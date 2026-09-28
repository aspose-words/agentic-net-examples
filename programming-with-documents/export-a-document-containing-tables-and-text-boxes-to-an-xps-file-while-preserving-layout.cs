using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;   // Needed for Table class

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a table with two rows and two columns.
        Table table = builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Write("Cell 3");
        builder.InsertCell();
        builder.Write("Cell 4");
        builder.EndRow();

        builder.EndTable();

        // Add a line break after the table.
        builder.Writeln();

        // Create a text box (shape of type TextBox).
        Shape textBox = new Shape(doc, ShapeType.TextBox)
        {
            Width = 200,
            Height = 100,
            WrapType = WrapType.Inline
        };

        // Add text inside the text box.
        Paragraph tbParagraph = new Paragraph(doc);
        Run tbRun = new Run(doc, "This is a text box.");
        tbParagraph.AppendChild(tbRun);
        textBox.AppendChild(tbParagraph);

        // Insert the text box into the document.
        builder.InsertNode(textBox);

        // Define output path for the XPS file.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "OutputDocument.xps");

        // Save the document as XPS.
        doc.Save(outputPath, SaveFormat.Xps);

        // Verify that the XPS file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine("XPS file saved successfully: " + outputPath);
        }
        else
        {
            Console.WriteLine("Failed to save XPS file.");
        }
    }
}
