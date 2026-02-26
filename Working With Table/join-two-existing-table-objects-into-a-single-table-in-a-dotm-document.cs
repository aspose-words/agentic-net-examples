using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableJoinExample
{
    class Program
    {
        static void Main()
        {
            // Load the existing DOTM document that contains at least two tables.
            Document doc = new Document("InputDocument.dotm");

            // Retrieve the first table from the document's body.
            Table firstTable = doc.FirstSection.Body.Tables[0];

            // Retrieve the second table using the GetChild method (index 1).
            Table secondTable = (Table)doc.GetChild(NodeType.Table, 1, true);

            // Move all rows from the second table to the first table.
            while (secondTable.HasChildNodes)
                firstTable.Rows.Add(secondTable.FirstRow);

            // Remove the now‑empty second table from the document.
            secondTable.Remove();

            // Save the modified document.
            doc.Save("OutputDocument.dotm");
        }
    }
}
