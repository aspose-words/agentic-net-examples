using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableJoinExample
{
    class Program
    {
        static void Main()
        {
            // Load the existing DOTM document.
            Document doc = new Document("InputTemplate.dotm");

            // Retrieve the first table from the document's body collection.
            Table firstTable = doc.FirstSection.Body.Tables[0];

            // Retrieve the second table using the GetChild method.
            Table secondTable = (Table)doc.GetChild(NodeType.Table, 1, true);

            // Transfer all rows from the second table to the first table.
            while (secondTable.HasChildNodes)
                firstTable.Rows.Add(secondTable.FirstRow);

            // Remove the now empty second table container.
            secondTable.Remove();

            // Save the document with the combined table.
            doc.Save("CombinedTables.dotm");
        }
    }
}
