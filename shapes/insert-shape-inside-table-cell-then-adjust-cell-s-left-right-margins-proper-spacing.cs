using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

class InsertShapeInTableCell
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table and insert the first cell.
        builder.StartTable();
        builder.InsertCell();

        // Insert an inline shape (rectangle) into the current cell.
        // Inline shapes are automatically placed inside the paragraph of the cell.
        Shape shape = builder.InsertShape(ShapeType.Rectangle, 100, 50);
        // Ensure the shape is treated as part of the cell's layout.
        shape.IsLayoutInCell = true;

        // Optionally add some text after the shape inside the same cell.
        builder.Writeln("Shape inside cell.");

        // Adjust the left and right padding (margins) of the current cell for proper spacing.
        // Padding values are in points; 10 points ≈ 0.14 inch.
        builder.CellFormat.LeftPadding = 10;   // left margin
        builder.CellFormat.RightPadding = 10;  // right margin

        // End the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document to the file system.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "ShapeInTableCell.docx");
        doc.Save(outputPath);
    }
}
