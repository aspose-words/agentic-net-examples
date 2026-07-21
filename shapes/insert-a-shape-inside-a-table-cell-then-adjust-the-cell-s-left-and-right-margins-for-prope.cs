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

        // Start a table and insert the first cell.
        builder.StartTable();
        builder.InsertCell();

        // Insert a rectangle shape inside the current cell.
        Shape shape = builder.InsertShape(ShapeType.Rectangle, 100, 50);
        shape.FillColor = System.Drawing.Color.LightBlue;
        shape.StrokeColor = System.Drawing.Color.DarkBlue;

        // Adjust the left and right margins (padding) of the cell for proper spacing.
        // The builder is positioned inside the cell, so we can obtain the cell from the current paragraph.
        Cell currentCell = (Cell)builder.CurrentParagraph.ParentNode;
        currentCell.CellFormat.LeftPadding = 10;   // points
        currentCell.CellFormat.RightPadding = 10;  // points

        // Finish the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document to disk.
        string outputPath = "ShapeInTableCell.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not saved correctly.");

        // Indicate successful completion.
        Console.WriteLine("Document created: " + Path.GetFullPath(outputPath));
    }
}
