using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a floating textbox shape.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 300, 200);
        // Optional: set the textbox to be floating (default for TextBox shape).
        textBox.WrapType = WrapType.None;

        // Move the builder cursor into the textbox so that subsequent content is placed there.
        builder.MoveTo(textBox.FirstParagraph);

        // Build a simple 2x2 table inside the textbox.
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

        // Finish the table.
        builder.EndTable();

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "FloatingTextboxTable.docx");
        doc.Save(outputPath);
    }
}
