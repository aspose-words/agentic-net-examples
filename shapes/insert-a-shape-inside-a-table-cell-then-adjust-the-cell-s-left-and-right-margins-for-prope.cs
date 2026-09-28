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
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table with a single cell.
        builder.StartTable();
        builder.InsertCell();

        // Insert a rectangle shape inside the current cell.
        Shape shape = builder.InsertShape(ShapeType.Rectangle, 100, 50);
        shape.FillColor = Color.LightBlue;
        shape.StrokeColor = Color.DarkBlue;

        // Retrieve the cell that contains the shape.
        Cell cell = builder.CurrentParagraph.ParentNode as Cell;
        if (cell == null)
            throw new Exception("Current node is not a table cell.");

        // Adjust left and right padding (margins) of the cell.
        cell.CellFormat.LeftPadding = 10;   // points
        cell.CellFormat.RightPadding = 10; // points

        // Complete the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document.
        string outputPath = "ShapeInTableCell.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("Document was not saved successfully.");

        Console.WriteLine($"Document saved to: {Path.GetFullPath(outputPath)}");
    }
}
