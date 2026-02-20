using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 2‑column table with one initial row.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1, Row 1");
        builder.InsertCell();
        builder.Write("Cell 2, Row 1");
        builder.EndRow();
        builder.EndTable();

        // Create a new row.
        Row newRow = new Row(doc);

        // First cell of the new row.
        Cell cell1 = new Cell(doc);
        cell1.AppendChild(new Paragraph(doc));
        cell1.FirstParagraph.AppendChild(new Run(doc, "Cell 1, Row 2"));
        newRow.Cells.Add(cell1);

        // Second cell of the new row.
        Cell cell2 = new Cell(doc);
        cell2.AppendChild(new Paragraph(doc));
        cell2.FirstParagraph.AppendChild(new Run(doc, "Cell 2, Row 2"));
        newRow.Cells.Add(cell2);

        // Add the new row to the table.
        table.Rows.Add(newRow);

        // Save the document.
        doc.Save("AddedRow.docx");
    }
}
