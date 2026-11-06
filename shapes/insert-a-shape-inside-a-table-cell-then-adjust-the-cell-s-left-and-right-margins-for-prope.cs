using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Output file path
        string outputPath = "ShapeInTableCell.docx";

        // Create a new document and a DocumentBuilder
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table with a single cell
        builder.StartTable();
        builder.InsertCell();

        // Insert a rectangle shape into the current cell
        builder.InsertShape(ShapeType.Rectangle, 100, 50);

        // Retrieve the cell that now contains the shape
        Cell cell = (Cell)builder.CurrentParagraph.ParentNode;

        // Adjust left and right margins (padding) of the cell
        cell.CellFormat.LeftPadding = 5;   // points
        cell.CellFormat.RightPadding = 5;  // points

        // Complete the table
        builder.EndRow();
        builder.EndTable();

        // Save the document
        doc.Save(outputPath);

        // Verify that the file was created
        if (!File.Exists(outputPath))
            throw new Exception("Failed to save the document.");
    }
}
