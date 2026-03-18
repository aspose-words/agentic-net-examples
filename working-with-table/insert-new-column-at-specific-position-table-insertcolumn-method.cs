using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 2x2 table.
        Table table = builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Write("R1C1");
        builder.InsertCell();
        builder.Write("R1C2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Write("R2C1");
        builder.InsertCell();
        builder.Write("R2C2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // -----------------------------------------------------------------
        // Insert a new column at position 1 (between the existing columns).
        // The Table.InsertColumn method is not available in older Aspose.Words
        // versions, so we add a new cell to each row manually.
        // -----------------------------------------------------------------
        int insertIndex = 1; // zero‑based column index where the new column will appear
        foreach (Row row in table.Rows)
        {
            // Insert a new cell at the desired position.
            Cell newCell = new Cell(doc);
            // Ensure the cell has at least one paragraph so we can add text later.
            newCell.AppendChild(new Paragraph(doc));
            row.Cells.Insert(insertIndex, newCell);
        }

        // Populate the newly inserted column with some text.
        table.Rows[0].Cells[insertIndex].FirstParagraph.AppendChild(new Run(doc, "New"));
        table.Rows[1].Cells[insertIndex].FirstParagraph.AppendChild(new Run(doc, "New"));

        // Save the document.
        doc.Save("InsertColumn.docx");
    }
}
