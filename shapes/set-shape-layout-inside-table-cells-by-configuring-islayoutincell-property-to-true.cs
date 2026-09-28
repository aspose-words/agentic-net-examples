using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table with a single cell.
        builder.StartTable();
        builder.InsertCell();

        // Insert some text before the shape.
        builder.Write("Before shape ");

        // Insert a rectangle shape into the cell.
        Shape shape = builder.InsertShape(ShapeType.Rectangle, 100, 50);
        // Configure the shape to layout inside the table cell.
        shape.IsLayoutInCell = true;

        // Insert some text after the shape.
        builder.Write(" After shape");

        // End the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document.
        string outputPath = "ShapeInTable.docx";
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception($"Failed to create the output file: {outputPath}");
    }
}
