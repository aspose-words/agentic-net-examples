using System;
using Aspose.Words;
using Aspose.Words.Tables;

class TableCombiner
{
    static void Main()
    {
        // Load the document that contains the two tables.
        Document doc = new Document("Input.docx");

        // Get the first table from the document's body.
        Table firstTable = doc.FirstSection.Body.Tables[0];

        // Get the second table using the GetChild method (index 1 = second table).
        Table secondTable = (Table)doc.GetChild(NodeType.Table, 1, true);

        // Move all rows from the second table to the first table.
        while (secondTable.HasChildNodes)
        {
            // Add the first row of the second table to the first table.
            firstTable.Rows.Add(secondTable.FirstRow);
        }

        // Remove the now empty second table from the document.
        secondTable.Remove();

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
