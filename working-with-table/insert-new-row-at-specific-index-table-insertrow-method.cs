using System;
using Aspose.Words;
using Aspose.Words.Tables;

class InsertRowAtSpecificIndex
{
    static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 2‑row table.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Row 0, Cell 0");
        builder.InsertCell();
        builder.Write("Row 0, Cell 1");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("Row 1, Cell 0");
        builder.InsertCell();
        builder.Write("Row 1, Cell 1");
        builder.EndRow();

        builder.EndTable();

        // ------------------------------------------------------------
        // Insert a new row at index 1 (between the two existing rows).
        // ------------------------------------------------------------

        // Create a new Row belonging to the same document.
        Row newRow = new Row(doc);
        // Ensure the row has at least one cell so it is valid.
        newRow.EnsureMinimum();

        // Populate the cells of the new row.
        // The first cell.
        Cell cell0 = new Cell(doc);
        cell0.AppendChild(new Paragraph(doc));
        cell0.FirstParagraph.AppendChild(new Run(doc, "Inserted Row, Cell 0"));
        newRow.AppendChild(cell0);

        // The second cell.
        Cell cell1 = new Cell(doc);
        cell1.AppendChild(new Paragraph(doc));
        cell1.FirstParagraph.AppendChild(new Run(doc, "Inserted Row, Cell 1"));
        newRow.AppendChild(cell1);

        // Insert the row into the table at the desired index.
        // RowCollection implements NodeCollection, so we can use Insert.
        table.Rows.Insert(1, newRow);

        // Save the document.
        doc.Save("InsertedRow.docx");
    }
}
