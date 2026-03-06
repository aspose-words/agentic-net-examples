using System;
using Aspose.Words;
using Aspose.Words.Tables;

class TableJoiner
{
    static void Main()
    {
        // Load the DOCM document that contains the tables.
        Document doc = new Document("input.docm");

        // Retrieve the first two tables in the document.
        // Adjust the indices if the tables are located elsewhere.
        Table firstTable = doc.FirstSection.Body.Tables[0];
        Table secondTable = (Table)doc.GetChild(NodeType.Table, 1, true);

        // Move all rows from the second table to the first table.
        while (secondTable.HasChildNodes)
        {
            // Add the first row of the second table to the first table.
            firstTable.Rows.Add(secondTable.FirstRow);
        }

        // Remove the now‑empty second table from the document.
        secondTable.Remove();

        // Save the modified document.
        doc.Save("output.docm");
    }
}
