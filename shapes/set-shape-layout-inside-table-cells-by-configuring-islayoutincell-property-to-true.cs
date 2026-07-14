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

        // Build a simple 1x1 table.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell with shape");
        builder.EndTable();

        // Move the cursor to the first paragraph of the first cell.
        Table table = doc.FirstSection.Body.Tables[0];
        builder.MoveTo(table.FirstRow.FirstCell.FirstParagraph);

        // Insert a floating rectangle shape.
        Shape shape = builder.InsertShape(
            ShapeType.Rectangle,
            RelativeHorizontalPosition.LeftMargin, 50,
            RelativeVerticalPosition.TopMargin, 100,
            100, 100,
            WrapType.None);

        // Configure the shape to be laid out inside the table cell.
        shape.IsLayoutInCell = true;
        shape.WrapType = WrapType.None; // Required for IsLayoutInCell to take effect.

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ShapeLayoutInCell.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The document was not saved successfully.");
    }
}
