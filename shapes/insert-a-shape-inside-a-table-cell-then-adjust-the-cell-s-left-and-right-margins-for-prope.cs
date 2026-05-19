using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class ShapeInTableCellExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table and insert the first cell.
        builder.StartTable();
        builder.InsertCell();

        // Adjust the left and right margins (padding) of the current cell.
        // Values are in points (1 point = 1/72 inch).
        builder.CellFormat.LeftPadding = 10;   // 10 points left margin
        builder.CellFormat.RightPadding = 10;  // 10 points right margin

        // Insert a simple rectangle shape inside the cell.
        Shape shape = builder.InsertShape(ShapeType.Rectangle, 50, 50);
        shape.FillColor = System.Drawing.Color.LightBlue;
        shape.StrokeColor = System.Drawing.Color.DarkBlue;

        // Finish the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Define the output file name.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ShapeInTableCell.docx");

        // Save the document.
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }
    }
}
