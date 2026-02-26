using System;
using Aspose.Words;
using Aspose.Words.Tables;

class JoinTablesInMhtml
{
    static void Main()
    {
        // Load the source MHTML document that already contains two tables.
        string inputFile = "input.mhtml";
        Document doc = new Document(inputFile);

        // Get the first table from the document (index 0 in the Tables collection).
        Table firstTable = doc.FirstSection.Body.Tables[0];

        // Get the second table using the GetChild method (the second table has index 1).
        Table secondTable = (Table)doc.GetChild(NodeType.Table, 1, true);

        // Move all rows from the second table to the first table.
        while (secondTable.HasChildNodes)
        {
            // Append the first row of the second table to the first table.
            firstTable.Rows.Add(secondTable.FirstRow);
        }

        // Remove the now‑empty second table container.
        secondTable.Remove();

        // Save the modified document back to MHTML format.
        string outputFile = "output.mhtml";
        doc.Save(outputFile, SaveFormat.Mhtml);
    }
}
