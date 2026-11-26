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
        textBox.WrapType = WrapType.None;
        textBox.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        textBox.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        textBox.Left = 100;
        textBox.Top = 100;

        // Move the builder into the textbox's first paragraph.
        builder.MoveTo(textBox.FirstParagraph);

        // Build a simple 2x2 table inside the textbox.
        builder.StartTable();

        // First row, first cell.
        builder.InsertCell();
        builder.Write("Cell 1");

        // First row, second cell.
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();

        // Second row, first cell.
        builder.InsertCell();
        builder.Write("Cell 3");

        // Second row, second cell.
        builder.InsertCell();
        builder.Write("Cell 4");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Save the document.
        string outputPath = "FloatingTextboxTable.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception($"Failed to create the output file: {outputPath}");
        }

        // Optionally, inform that the process completed.
        Console.WriteLine($"Document saved successfully to '{outputPath}'.");
    }
}
