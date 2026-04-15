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

        // Insert content using the macro‑like routine.
        InsertTableParagraphAndLinkedTextBox(builder);

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "MacroInsert.docx");
        doc.Save(outputPath);
    }

    // Inserts a table, a paragraph, and a linked text box at the current cursor position.
    private static void InsertTableParagraphAndLinkedTextBox(DocumentBuilder builder)
    {
        // ---------- Insert a 2x2 table ----------
        builder.StartTable();

        // First row
        builder.InsertCell();
        builder.Write("Cell 1,1");
        builder.InsertCell();
        builder.Write("Cell 1,2");
        builder.EndRow();

        // Second row
        builder.InsertCell();
        builder.Write("Cell 2,1");
        builder.InsertCell();
        builder.Write("Cell 2,2");
        builder.EndRow();

        builder.EndTable();

        // Add a paragraph after the table.
        builder.Writeln(); // Ensure we are on a new line.
        builder.Writeln("This paragraph follows the table.");

        // ---------- Insert a linked (floating) text box ----------
        // Create a floating text box shape.
        Shape textBox = new Shape(builder.Document, ShapeType.TextBox);
        textBox.WrapType = WrapType.None;
        textBox.Height = 50;
        textBox.Width = 200;
        textBox.HorizontalAlignment = HorizontalAlignment.Center;
        textBox.VerticalAlignment = VerticalAlignment.Top;

        // Add a paragraph with text inside the text box.
        textBox.AppendChild(new Paragraph(builder.Document));
        Paragraph para = textBox.FirstParagraph;
        para.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        Run run = new Run(builder.Document, "Linked Text Box Content");
        para.AppendChild(run);

        // Insert the text box at the current position.
        builder.InsertNode(textBox);
    }
}
