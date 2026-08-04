using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;   // Needed for the Table class

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 2‑cell table.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("First cell");
        builder.InsertCell();
        builder.Write("Second cell");
        builder.EndRow();
        builder.EndTable();

        // Move the cursor to the first cell where the shape will be placed.
        builder.MoveTo(table.FirstRow.FirstCell.FirstParagraph);

        // Insert a floating rectangle shape.
        Shape shape = builder.InsertShape(
            ShapeType.Rectangle,
            RelativeHorizontalPosition.LeftMargin, 0,
            RelativeVerticalPosition.TopMargin, 0,
            100, 100,
            WrapType.None);

        // Configure the shape to be laid out inside the table cell.
        shape.IsLayoutInCell = true;

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ShapeLayoutInCell.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The document was not saved correctly.");
    }
}
