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

        // Perform the macro‑like operation.
        InsertTableParagraphAndLinkedTextBox(doc);

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "MacroOutput.docx");
        doc.Save(outputPath);

        // Verify that the file was created (non‑interactive).
        if (File.Exists(outputPath))
        {
            // Reopen to ensure it can be loaded without error.
            Document loaded = new Document(outputPath);
        }
    }

    private static void InsertTableParagraphAndLinkedTextBox(Document doc)
    {
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a simple 2x2 table.
        builder.StartTable();

        builder.InsertCell();
        builder.Writeln("Cell 1,1");
        builder.InsertCell();
        builder.Writeln("Cell 1,2");
        builder.EndRow();

        builder.InsertCell();
        builder.Writeln("Cell 2,1");
        builder.InsertCell();
        builder.Writeln("Cell 2,2");
        builder.EndRow();

        builder.EndTable();

        // Insert a paragraph after the table.
        builder.Writeln("This paragraph follows the table.");

        // Insert a linked text box.
        Shape textBox = new Shape(doc, ShapeType.TextBox);
        textBox.Width = 200;
        textBox.Height = 100;
        textBox.WrapType = WrapType.Inline;
        textBox.HRef = "https://www.example.com";

        // Add text to the text box.
        Paragraph para = new Paragraph(doc);
        Run run = new Run(doc, "Click to visit example.com");
        para.AppendChild(run);
        textBox.AppendChild(para);

        // Insert the text box into the document.
        builder.InsertNode(textBox);
    }
}
