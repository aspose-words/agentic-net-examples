using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load the HTML document that contains the tables.
        Document doc = new Document("input.html");

        // Retrieve the first table from the document's body.
        Table firstTable = doc.FirstSection.Body.Tables[0];

        // Retrieve the second table using the GetChild method.
        Table secondTable = (Table)doc.GetChild(NodeType.Table, 1, true);

        // Move all rows from the second table to the first table.
        while (secondTable.HasChildNodes)
            firstTable.Rows.Add(secondTable.FirstRow);

        // Remove the now‑empty second table container.
        secondTable.Remove();

        // Save the modified document back to HTML (or any other supported format).
        doc.Save("output.html");
    }
}
