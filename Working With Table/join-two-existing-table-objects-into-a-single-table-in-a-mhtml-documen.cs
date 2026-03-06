using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load the source MHTML document that contains the two tables.
        Document doc = new Document("input.mht");

        // Retrieve the first table from the document's body.
        Table firstTable = doc.FirstSection.Body.Tables[0];

        // Retrieve the second table using the generic GetChild method.
        Table secondTable = (Table)doc.GetChild(NodeType.Table, 1, true);

        // Transfer all rows from the second table to the first table.
        // The Rows.Add method automatically removes the moved row from its original table.
        while (secondTable.HasChildNodes)
        {
            firstTable.Rows.Add(secondTable.FirstRow);
        }

        // After all rows have been moved, remove the now‑empty second table node.
        secondTable.Remove();

        // Save the modified document back to MHTML format.
        doc.Save("output.mht");
    }
}
