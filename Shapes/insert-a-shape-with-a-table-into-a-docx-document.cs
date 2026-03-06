using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;
using System.Drawing;

class InsertShapeWithTable
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a DocumentBuilder which will be used to insert content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a floating text box shape. Width = 200 points, Height = 100 points.
        Shape shape = builder.InsertShape(ShapeType.TextBox, 200, 100);
        // Make the shape float (no text wrapping) so it behaves like a container.
        shape.WrapType = WrapType.None;

        // -------------------------
        // Build a simple 1x1 table that will be placed inside the shape.
        // -------------------------

        // Create a new table that belongs to the same document.
        Table table = new Table(doc);
        // Ensure the table has at least one row, cell and paragraph.
        table.EnsureMinimum();

        // Get the first (and only) cell.
        Cell cell = table.FirstRow.FirstCell;
        // Remove the default empty paragraph.
        cell.RemoveAllChildren();
        // Add a paragraph with some text to the cell.
        Paragraph para = new Paragraph(doc);
        para.AppendChild(new Run(doc, "Table inside a shape"));
        cell.AppendChild(para);

        // Append the table to the shape. Shapes can contain tables as child nodes.
        shape.AppendChild(table);

        // Save the document to a DOCX file.
        doc.Save("ShapeWithTable.docx");
    }
}
