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

        // Insert a floating textbox shape.
        Shape textbox = new Shape(doc, ShapeType.TextBox);
        textbox.Width = 200;
        textbox.Height = 100;
        textbox.WrapType = WrapType.Square;
        textbox.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        textbox.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        textbox.HorizontalAlignment = HorizontalAlignment.Right;
        textbox.VerticalAlignment = VerticalAlignment.Top;
        builder.InsertNode(textbox);

        // Add a paragraph after the shape to position the table.
        builder.Writeln();

        // Start a table.
        builder.StartTable();

        // First row, first cell.
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.EndRow();

        // Second row, first cell.
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document.
        string fileName = "TableOverlap.docx";
        doc.Save(fileName);

        // Verify that the file was saved.
        if (!File.Exists(fileName))
            throw new Exception("Document was not saved.");

        Console.WriteLine($"Document saved successfully: {fileName}");
    }
}
