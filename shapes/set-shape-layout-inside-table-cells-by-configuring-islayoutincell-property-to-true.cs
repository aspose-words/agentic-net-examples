using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;   // Needed for the Table class

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table with a single cell.
        Table table = builder.StartTable();
        builder.InsertCell();

        // Move the cursor to the first paragraph of the first cell.
        builder.MoveTo(table.FirstRow.FirstCell.FirstParagraph);

        // Insert a floating rectangle shape inside the cell.
        Shape shape = builder.InsertShape(
            ShapeType.Rectangle,
            RelativeHorizontalPosition.LeftMargin, 50,
            RelativeVerticalPosition.TopMargin, 100,
            100, 100,
            WrapType.None);

        // Configure the shape to be laid out inside the table cell.
        shape.IsLayoutInCell = true;
        shape.WrapType = WrapType.None; // Required for IsLayoutInCell to take effect.

        // End the table.
        builder.EndTable();

        // Define the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ShapeLayoutInCell.docx");

        // Save the document.
        doc.Save(outputPath);

        // Validate that the file was created and the property is set correctly.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not created.");

        if (!shape.IsLayoutInCell)
            throw new InvalidOperationException("IsLayoutInCell property was not set to true.");
    }
}
