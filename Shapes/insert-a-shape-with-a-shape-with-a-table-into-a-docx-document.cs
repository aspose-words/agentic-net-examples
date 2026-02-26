using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a floating text box shape that will hold the table.
        Shape shape = builder.InsertShape(ShapeType.TextBox, 300, 200);
        shape.WrapType = WrapType.None;               // No text wrapping.
        shape.IsLayoutInCell = false;                 // Ensure the shape is not treated as a table cell.

        // Build a simple table.
        Table table = new Table(doc);
        table.EnsureMinimum();                        // Guarantees at least one row, cell, and paragraph.

        // Add some content to the first cell.
        Cell firstCell = table.FirstRow.FirstCell;
        firstCell.FirstParagraph.AppendChild(new Run(doc, "Table inside a shape"));

        // Append the table to the shape's child nodes.
        shape.AppendChild(table);

        // Move the cursor to the end of the document (optional).
        builder.MoveToDocumentEnd();

        // Save the resulting document.
        doc.Save("ShapeWithTable.docx");
    }
}
