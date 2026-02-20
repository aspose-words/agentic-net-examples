using System;
using Aspose.Words;
using Aspose.Words.Tables;

class JoinTablesExample
{
    static void Main()
    {
        // Load the DOTX template that already contains the two tables.
        Document doc = new Document("Template.dotx");

        // Retrieve the first two tables from the document.
        // Adjust the indexes if the tables are located elsewhere.
        Table firstTable = doc.FirstSection.Body.Tables[0];
        Table secondTable = doc.FirstSection.Body.Tables[1];

        // Move all rows from the second table into the first table.
        // Use Document.ImportNode to import each row into the destination document.
        foreach (Row row in secondTable.Rows)
        {
            // Import the row into the document's context (deep clone).
            Row importedRow = (Row)doc.ImportNode(row, true);
            // Append the imported row to the first table.
            firstTable.Rows.Add(importedRow);
        }

        // Remove the now‑empty second table from the document.
        secondTable.Remove();

        // Save the modified document back to DOTX format.
        doc.Save("JoinedTables.dotx");
    }
}
