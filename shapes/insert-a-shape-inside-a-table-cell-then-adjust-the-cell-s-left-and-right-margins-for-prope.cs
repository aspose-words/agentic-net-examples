using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder to work with it.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table.
        builder.StartTable();

        // Insert the first cell where the shape will be placed.
        builder.InsertCell();

        // Adjust the left and right margins (paddings) of the current cell.
        // This creates spacing between the cell borders and its contents.
        builder.CellFormat.SetPaddings(leftPadding: 20, topPadding: 0, rightPadding: 20, bottomPadding: 0);

        // Insert a simple rectangle shape inside the cell.
        // Using InsertShape follows the preferred core rule for shape insertion.
        Shape shape = builder.InsertShape(ShapeType.Rectangle, width: 100, height: 50);
        shape.FillColor = System.Drawing.Color.LightBlue;
        shape.Stroke.Color = System.Drawing.Color.DarkBlue;

        // Optionally add some text after the shape inside the same cell.
        builder.Writeln();
        builder.Write("Shape inside a table cell with adjusted margins.");

        // End the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Define the output file path.
        string outputPath = "ShapeInTableCell.docx";

        // Save the document.
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
    }
}
