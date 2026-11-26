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

        // Start a table with a single row and two cells.
        builder.StartTable();

        // First cell – insert a shape.
        builder.InsertCell();
        Shape shape = builder.InsertShape(ShapeType.Rectangle, 100, 50);
        // Configure the shape to layout inside the table cell.
        shape.IsLayoutInCell = true;

        // Second cell – add some text.
        builder.InsertCell();
        builder.Writeln("Second cell");

        // End the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document to disk.
        string outputPath = "ShapeInTable.docx";
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }
    }
}
