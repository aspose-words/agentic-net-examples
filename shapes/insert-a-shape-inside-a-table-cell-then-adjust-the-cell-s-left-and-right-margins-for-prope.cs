using System;
using System.IO;
using System.Drawing;
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

        // Start a table with one row and two cells.
        builder.StartTable();

        // First cell (optional content).
        builder.InsertCell();
        builder.Writeln("First cell");

        // Second cell where the shape will be inserted.
        builder.InsertCell();

        // Insert a rectangle shape into the current cell.
        Shape shape = builder.InsertShape(ShapeType.Rectangle, 100, 50);
        shape.StrokeColor = Color.Black;
        shape.FillColor = Color.LightBlue;

        // End the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Retrieve the second cell (index 1) to adjust its margins.
        Cell targetCell = doc.GetChildNodes(NodeType.Cell, true)[1] as Cell;
        if (targetCell == null)
            throw new Exception("Target cell not found.");

        // Set left and right padding (cell margins) to 10 points.
        targetCell.CellFormat.LeftPadding = 10;
        targetCell.CellFormat.RightPadding = 10;

        // Save the document.
        string outputPath = "ShapeInTableCell.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception($"Failed to create output file: {outputPath}");
    }
}
