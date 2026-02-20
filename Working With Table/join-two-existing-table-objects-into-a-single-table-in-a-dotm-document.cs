using System;
using Aspose.Words;
using Aspose.Words.Tables;

class JoinTablesInDotm
{
    static void Main()
    {
        // Load the DOTM document.
        Document doc = new Document("InputTemplate.dotm");

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

        // Move each row from the second table to the end of the first table.
        // Import each row into the document of the first table to preserve ownership.
        foreach (Row row in secondTable.Rows)
        {
            // Import the row (deep copy) into the document.
            Row importedRow = (Row)firstTable.Document.ImportNode(row, true);
            // Add the imported row to the first table.
            firstTable.Rows.Add(importedRow);
        }

        // Remove the now-empty second table from the document.
        secondTable.Remove();

        // Save the modified document.
        doc.Save("OutputTemplate.dotm");
    }
}
