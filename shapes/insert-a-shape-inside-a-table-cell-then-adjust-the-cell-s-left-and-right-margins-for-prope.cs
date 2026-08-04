using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Begin a table.
        builder.StartTable();

        // Insert the first cell of the table.
        builder.InsertCell();

        // Insert a rectangle shape into the current cell.
        Shape shape = builder.InsertShape(ShapeType.Rectangle, 50, 50);
        shape.FillColor = Color.LightBlue;

        // Adjust the cell's left and right margins (padding) to give the shape space.
        builder.CellFormat.LeftPadding = 10;   // points
        builder.CellFormat.RightPadding = 10; // points

        // Optional text to illustrate the padding effect.
        builder.Writeln("Shape inside cell");

        // Finish the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ShapeInTableCell.docx");
        doc.Save(outputPath);

        // Validate that the file was saved.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Document was not saved successfully.");
    }
}
