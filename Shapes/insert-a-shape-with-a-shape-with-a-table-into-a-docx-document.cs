using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a floating text box shape that will contain the table.
        Shape shape = builder.InsertShape(ShapeType.TextBox, 200, 100);
        shape.WrapType = WrapType.None;               // No text wrapping.
        shape.IsLayoutInCell = false;                 // Ensure the shape is a top‑level object.

        // Build a simple 2×2 table.
        Table table = new Table(doc);
        table.EnsureMinimum(); // Guarantees at least one row, cell, and paragraph.

        // First row.
        Row row1 = new Row(doc);
        table.AppendChild(row1);

        Cell cell11 = new Cell(doc);
        cell11.AppendChild(new Paragraph(doc));
        cell11.FirstParagraph.AppendChild(new Run(doc, "Cell 1,1"));
        row1.AppendChild(cell11);

        Cell cell12 = new Cell(doc);
        cell12.AppendChild(new Paragraph(doc));
        cell12.FirstParagraph.AppendChild(new Run(doc, "Cell 1,2"));
        row1.AppendChild(cell12);

        // Second row.
        Row row2 = new Row(doc);
        table.AppendChild(row2);

        Cell cell21 = new Cell(doc);
        cell21.AppendChild(new Paragraph(doc));
        cell21.FirstParagraph.AppendChild(new Run(doc, "Cell 2,1"));
        row2.AppendChild(cell21);

        Cell cell22 = new Cell(doc);
        cell22.AppendChild(new Paragraph(doc));
        cell22.FirstParagraph.AppendChild(new Run(doc, "Cell 2,2"));
        row2.AppendChild(cell22);

        // Insert the table into the shape.
        shape.AppendChild(table);

        // Save the resulting document.
        doc.Save("ShapeWithTable.docx");
    }
}
