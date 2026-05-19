using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

public class ShapeLayoutInCellExample
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table and add a single cell.
        Table table = builder.StartTable();
        builder.InsertCell();

        // Move the cursor to the first paragraph of the cell.
        builder.MoveTo(table.FirstRow.FirstCell.FirstParagraph);

        // Insert a floating rectangle shape inside the cell.
        // Use RelativeHorizontalPosition.LeftMargin and RelativeVerticalPosition.TopMargin as the origin.
        Shape shape = builder.InsertShape(
            ShapeType.Rectangle,
            RelativeHorizontalPosition.LeftMargin, 50,
            RelativeVerticalPosition.TopMargin, 20,
            100, 50,
            WrapType.None);

        // Configure the shape to be laid out inside the table cell.
        shape.IsLayoutInCell = true;
        shape.WrapType = WrapType.None; // Required for IsLayoutInCell to take effect.

        // End the table.
        builder.EndTable();

        // Save the document.
        string fileName = "ShapeLayoutInCell.docx";
        doc.Save(fileName);

        // Validate that the file was created.
        if (!File.Exists(fileName))
            throw new InvalidOperationException("The document was not saved correctly.");

        // Validate that the shape property is set as expected.
        Shape savedShape = (Shape)doc.GetChild(NodeType.Shape, 0, true);
        if (!savedShape.IsLayoutInCell)
            throw new InvalidOperationException("IsLayoutInCell property was not set to true.");
    }
}
