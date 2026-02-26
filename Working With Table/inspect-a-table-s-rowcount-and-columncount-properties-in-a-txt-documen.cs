using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Loading;

class Program
{
    static void Main()
    {
        // Load the TXT document using default TxtLoadOptions.
        Document doc = new Document("input.txt", new TxtLoadOptions());

        // Iterate through all tables in the document.
        foreach (Table table in doc.GetChildNodes(NodeType.Table, true))
        {
            // Number of rows in the table.
            int rowCount = table.Rows.Count;

            // Number of columns in the table (based on the first row).
            int columnCount = table.FirstRow != null ? table.FirstRow.Cells.Count : 0;

            Console.WriteLine($"Table found: {rowCount} rows x {columnCount} columns");
        }

        // Optionally save the document after inspection.
        doc.Save("output.docx");
    }
}
