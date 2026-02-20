using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Loading;

class TableInspection
{
    static void Main()
    {
        // Load a plain‑text document. TxtLoadOptions allows us to specify any additional
        // loading behavior for TXT files (e.g., encoding, direction, etc.).
        TxtLoadOptions loadOptions = new TxtLoadOptions();
        Document doc = new Document("Input.txt", loadOptions);

        // Access the collection of tables in the first section's body.
        TableCollection tables = doc.FirstSection.Body.Tables;

        // Iterate through each table and output its row and column counts.
        for (int i = 0; i < tables.Count; i++)
        {
            Table table = tables[i];

            // Row count is directly available.
            int rowCount = table.Rows.Count;

            // Column count is derived from the number of cells in the first row.
            // (All rows in a well‑formed table have the same number of cells.)
            int columnCount = table.Rows[0].Cells.Count;

            Console.WriteLine($"Table {i}: Rows = {rowCount}, Columns = {columnCount}");
        }

        // (Optional) Save the document if any modifications were made.
        // doc.Save("Output.docx");
    }
}
