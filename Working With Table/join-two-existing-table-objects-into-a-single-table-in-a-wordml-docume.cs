using System;
using Aspose.Words;
using Aspose.Words.Tables;

class TableMerger
{
    static void Main()
    {
        // Load the existing Word document that contains at least two tables.
        Document doc = new Document("input.docx");

        // Retrieve the first table from the document's body.
        Table firstTable = doc.FirstSection.Body.Tables[0];

        // Retrieve the second table using the GetChild method (index 1 for the second table).
        Table secondTable = (Table)doc.GetChild(NodeType.Table, 1, true);

        // Move all rows from the second table to the first table.
        // While the second table still has rows, add its first row to the first table.
        while (secondTable.HasChildNodes)
        {
            // Add the first row of the second table to the first table.
            firstTable.Rows.Add(secondTable.FirstRow);
        }

        // Remove the now-empty second table from the document.
        secondTable.Remove();

        // Save the modified document to a new file.
        doc.Save("output.docx");
    }
}
