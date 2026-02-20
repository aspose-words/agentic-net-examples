using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Loading;

class Program
{
    static void Main()
    {
        // Load the RTF document.
        var loadOptions = new RtfLoadOptions();
        Document doc = new Document("input.rtf", loadOptions);

        // Access the collection of tables in the first section.
        TableCollection tables = doc.FirstSection.Body.Tables;

        // Iterate through each table and output its row and column counts.
        for (int i = 0; i < tables.Count; i++)
        {
            Table table = tables[i];
            int rowCount = table.Rows.Count;
            int columnCount = rowCount > 0 ? table.Rows[0].Cells.Count : 0;

            Console.WriteLine($"Table {i}: Rows = {rowCount}, Columns = {columnCount}");
        }
    }
}
