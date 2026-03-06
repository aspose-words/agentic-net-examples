using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load a document that already contains at least two tables.
        Document doc = new Document("Tables.docx");

        // Retrieve the first table from the document's body collection.
        Table firstTable = doc.FirstSection.Body.Tables[0];

        // Retrieve the second table using the generic GetChild method.
        Table secondTable = (Table)doc.GetChild(NodeType.Table, 1, true);

        // Transfer all rows from the second table to the first table.
        while (secondTable.HasChildNodes)
        {
            // The FirstRow property returns the current first row of the second table.
            // Adding it to firstTable.Rows automatically removes it from secondTable.
            firstTable.Rows.Add(secondTable.FirstRow);
        }

        // After moving all rows, the second table is empty; remove it from the document.
        secondTable.Remove();

        // Save the resulting document with the combined table.
        doc.Save("Table.CombineTables.docx");
    }
}
