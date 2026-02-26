using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load the source RTF document.
        Document doc = new Document("Input.rtf");

        // Retrieve the first two tables in the document.
        // The Body.Tables collection provides direct access to tables.
        Table firstTable = doc.FirstSection.Body.Tables[0];
        Table secondTable = doc.FirstSection.Body.Tables[1];

        // Transfer all rows from the second table to the first table.
        // While the second table still contains rows, move its first row.
        while (secondTable.HasChildNodes)
        {
            // Extract the first row from the second table.
            Row rowToMove = secondTable.FirstRow;
            // Detach the row from its current parent.
            secondTable.RemoveChild(rowToMove);
            // Append the row to the first table.
            firstTable.Rows.Add(rowToMove);
        }

        // Remove the now empty second table from the document.
        secondTable.Remove();

        // Save the modified document back to RTF format.
        doc.Save("Output.rtf");
    }
}
