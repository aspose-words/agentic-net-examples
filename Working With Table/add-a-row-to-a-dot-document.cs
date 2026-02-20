using System;
using Aspose.Words;
using Aspose.Words.Tables;

class AddRowExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to start a table with one initial row.
        DocumentBuilder builder = new DocumentBuilder(doc);
        Table table = builder.StartTable();

        // First row – add two cells with sample text.
        builder.InsertCell();
        builder.Write("Row 1, Cell 1");
        builder.InsertCell();
        builder.Write("Row 1, Cell 2");
        builder.EndRow();

        // End the table construction for now.
        builder.EndTable();

        // Create a new row instance associated with the same document.
        Row newRow = new Row(doc);

        // Add the required number of cells to the new row.
        // Each cell must be added to the row's Cells collection.
        Cell cell1 = new Cell(doc);
        cell1.AppendChild(new Paragraph(doc));
        cell1.FirstParagraph.AppendChild(new Run(doc, "Row 2, Cell 1"));
        newRow.Cells.Add(cell1);

        Cell cell2 = new Cell(doc);
        cell2.AppendChild(new Paragraph(doc));
        cell2.FirstParagraph.AppendChild(new Run(doc, "Row 2, Cell 2"));
        newRow.Cells.Add(cell2);

        // Insert the new row at the end of the table.
        table.Rows.Add(newRow);

        // Save the document to disk.
        doc.Save("AddedRow.docx");
    }
}
