using System;
using Aspose.Words;
using Aspose.Words.Tables;

class JoinTablesInHtml
{
    static void Main()
    {
        // Load the HTML document that already contains two tables.
        Document doc = new Document("input.html");

        // Access the collection of tables in the first section's body.
        TableCollection tables = doc.FirstSection.Body.Tables;

        // Ensure there are at least two tables to join.
        if (tables.Count < 2)
        {
            Console.WriteLine("The document does not contain two tables to join.");
            return;
        }

        // Get references to the first and second tables.
        Table firstTable = tables[0];
        Table secondTable = tables[1];

        // Append each row from the second table to the first table.
        // ImportNode is required because the rows belong to a different parent node.
        foreach (Row row in secondTable.Rows)
        {
            Row importedRow = (Row)firstTable.Document.ImportNode(row, true, ImportFormatMode.KeepSourceFormatting);
            firstTable.Rows.Add(importedRow);
        }

        // Remove the now redundant second table from the document.
        secondTable.Remove();

        // Save the modified document back to HTML.
        doc.Save("output.html");
    }
}
