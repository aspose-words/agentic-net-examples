using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Loading;

class Program
{
    static void Main()
    {
        // Load a plain‑text document. TxtLoadOptions allows us to specify TXT‑specific settings.
        var loadOptions = new TxtLoadOptions();
        Document doc = new Document("input.txt", loadOptions);

        // Iterate through all tables in the document.
        foreach (Table table in doc.FirstSection.Body.Tables)
        {
            // Row count is the number of Row objects in the table.
            int rowCount = table.Rows.Count;

            // Column count is the number of cells in the first row (if the table has at least one row).
            int columnCount = table.FirstRow != null ? table.FirstRow.Count : 0;

            Console.WriteLine($"Table found: {rowCount} rows x {columnCount} columns");
        }

        // Save the document (required by the lifecycle rule).
        doc.Save("output.docx");
    }
}
