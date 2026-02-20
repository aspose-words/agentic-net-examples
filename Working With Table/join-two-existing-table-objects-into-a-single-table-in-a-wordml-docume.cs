using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load the existing WordML document.
        Document doc = new Document("Input.docx");

        // Access the collection of tables in the first section's body.
        TableCollection tables = doc.FirstSection.Body.Tables;

        // Ensure there are at least two tables to join.
        if (tables.Count >= 2)
        {
            // Get references to the first and second tables.
            Table firstTable = tables[0];
            Table secondTable = tables[1];

            // Append each row from the second table to the first table.
            // Use Document.ImportNode (not Table.ImportNode) to clone the row into the same document.
            foreach (Row row in secondTable.Rows)
            {
                Row importedRow = (Row)doc.ImportNode(row, true);
                firstTable.Rows.Add(importedRow);
            }

            // Remove the now redundant second table from the document.
            secondTable.Remove();
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
