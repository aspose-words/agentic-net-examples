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

        // Start a table.
        builder.StartTable();

        // Insert the first cell where the shape will be placed.
        builder.InsertCell();

        // Insert a simple rectangle shape inside the current cell.
        Shape shape = builder.InsertShape(ShapeType.Rectangle, 50, 30);
        shape.FillColor = System.Drawing.Color.LightBlue;
        shape.StrokeColor = System.Drawing.Color.DarkBlue;

        // Retrieve the cell that currently contains the shape.
        Cell currentCell = (Cell)builder.CurrentParagraph.GetAncestor(NodeType.Cell);

        // Adjust the left and right margins (padding) of the cell for proper spacing.
        // Values are in points; 10 points ≈ 0.14 inch.
        currentCell.CellFormat.LeftPadding = 10;
        currentCell.CellFormat.RightPadding = 10;

        // Add a second cell to complete the row (optional content).
        builder.InsertCell();
        builder.Writeln("Second cell");

        // End the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Define the output file path.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "ShapeInTable.docx");

        // Save the document.
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The document was not saved correctly.");

        // The program finishes without waiting for user input.
    }
}
