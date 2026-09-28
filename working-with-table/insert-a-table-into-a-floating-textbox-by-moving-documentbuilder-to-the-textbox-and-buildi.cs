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

        // Insert a floating textbox shape.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 300, 200);
        // Ensure the shape is positioned relative to the page (optional).
        textBox.WrapType = WrapType.Inline;

        // Move the builder into the textbox's first paragraph.
        // The textbox contains its own paragraph collection.
        builder.MoveTo(textBox.FirstParagraph);

        // Build a simple 2x2 table inside the textbox.
        builder.StartTable();

        // First row
        builder.InsertCell();
        builder.Writeln("Cell 1,1");
        builder.InsertCell();
        builder.Writeln("Cell 1,2");
        builder.EndRow();

        // Second row
        builder.InsertCell();
        builder.Writeln("Cell 2,1");
        builder.InsertCell();
        builder.Writeln("Cell 2,2");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Save the document.
        string outputPath = "FloatingTextboxTable.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }
    }
}
