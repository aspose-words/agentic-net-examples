using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load the DOCX file that contains the two tables.
        Document doc = new Document("Tables.docx");

        // Get the first table from the body of the first section.
        Table firstTable = doc.FirstSection.Body.Tables[0];

        // Get the second table using the GetChild method (index 1 for the second table).
        Table secondTable = (Table)doc.GetChild(NodeType.Table, 1, true);

        // Transfer all rows from the second table to the first table.
        while (secondTable.HasChildNodes)
        {
            // Add the first row of the second table to the first table.
            firstTable.Rows.Add(secondTable.FirstRow);
        }

        // Remove the now empty second table from the document.
        secondTable.Remove();

        // Save the document with the combined table.
        doc.Save("CombinedTables.docx");
    }
}
