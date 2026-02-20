using System;
using Aspose.Words;
using Aspose.Words.Tables;

class JoinTablesInDocm
{
    static void Main()
    {
        // Load the existing DOCM document.
        Document doc = new Document("Input.docm");

        // Get the collection of tables in the first section's body.
        TableCollection tables = doc.FirstSection.Body.Tables;

        // Ensure there are at least two tables to join.
        if (tables.Count < 2)
        {
            Console.WriteLine("The document does not contain two tables to join.");
            return;
        }

        // Reference the first and second tables.
        Table firstTable = tables[0];
        Table secondTable = tables[1];

        // Append all rows from the second table to the first table.
        // Clone each row to keep the original formatting and content.
        foreach (Row row in secondTable.Rows)
        {
            // Clone the row (deep clone) and add it to the first table.
            Row clonedRow = (Row)row.Clone(true);
            firstTable.Rows.Add(clonedRow);
        }

        // Remove the now redundant second table from the document.
        secondTable.Remove();

        // Save the modified document as a DOCM file.
        doc.Save("Output.docm");
    }
}
