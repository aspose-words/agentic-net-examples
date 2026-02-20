using System;
using Aspose.Words;
using Aspose.Words.Tables;

class JoinTablesInMhtml
{
    static void Main()
    {
        // Load the MHTML document that contains the tables.
        Document doc = new Document("input.mht");

        // Get the collection of tables in the first section's body.
        TableCollection tables = doc.FirstSection.Body.Tables;

        // Ensure there are at least two tables to join.
        if (tables.Count < 2)
        {
            Console.WriteLine("The document must contain at least two tables.");
            return;
        }

        // Reference to the first and second tables.
        Table firstTable = tables[0];
        Table secondTable = tables[1];

        // Append each row from the second table to the first table.
        // Clone the rows to preserve all cell contents and formatting.
        foreach (Row row in secondTable.Rows)
        {
            // Deep clone the row (including all child nodes).
            Row clonedRow = (Row)row.Clone(true);
            firstTable.Rows.Add(clonedRow);
        }

        // Remove the now-empty second table from the document.
        secondTable.Remove();

        // Save the modified document back to MHTML format.
        doc.Save("output.mht");
    }
}
