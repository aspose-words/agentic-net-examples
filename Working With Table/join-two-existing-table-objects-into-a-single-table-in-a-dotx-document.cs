using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load the existing DOTX template that contains at least two tables.
        string inputPath = @"C:\Data\Template.dotx";
        Document doc = new Document(inputPath);

        // Retrieve the first table from the document's body.
        Table firstTable = doc.FirstSection.Body.Tables[0];

        // Retrieve the second table using the generic GetChild method.
        Table secondTable = (Table)doc.GetChild(NodeType.Table, 1, true);

        // Move all rows from the second table to the first table.
        while (secondTable.HasChildNodes)
        {
            // Add the first row of the second table to the first table.
            firstTable.Rows.Add(secondTable.FirstRow);
        }

        // Remove the now‑empty second table container.
        secondTable.Remove();

        // Save the modified document back to DOTX format.
        string outputPath = @"C:\Data\CombinedTable.dotx";
        doc.Save(outputPath);
    }
}
