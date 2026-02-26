using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load the document that already contains the two tables to be merged.
        Document doc = new Document("Tables.docx");

        // Retrieve the first table from the document's body.
        Table firstTable = doc.FirstSection.Body.Tables[0];

        // Retrieve the second table using the GetChild method (index 1 = second table).
        Table secondTable = (Table)doc.GetChild(NodeType.Table, 1, true);

        // Transfer all rows from the second table to the first table.
        // The loop continues until the second table has no more child nodes (rows).
        while (secondTable.HasChildNodes)
            firstTable.Rows.Add(secondTable.FirstRow);

        // Remove the now‑empty second table from the document.
        secondTable.Remove();

        // Save the modified document containing the combined table.
        doc.Save("CombinedTables.docx");
    }
}
