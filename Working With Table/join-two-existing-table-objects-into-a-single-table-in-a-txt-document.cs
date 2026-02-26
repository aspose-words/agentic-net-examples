using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Tables; // Added namespace for Table class

class Program
{
    static void Main()
    {
        // Load the document that contains the two tables to be merged.
        Document doc = new Document("Tables.docx");

        // Retrieve the first table from the document's body.
        Table firstTable = doc.FirstSection.Body.Tables[0];

        // Retrieve the second table using the GetChild method.
        Table secondTable = (Table)doc.GetChild(NodeType.Table, 1, true);

        // Move all rows from the second table to the first table.
        while (secondTable.HasChildNodes)
            firstTable.Rows.Add(secondTable.FirstRow);

        // Remove the now empty second table from the document.
        secondTable.Remove();

        // Configure TXT save options to preserve the visual layout of tables.
        TxtSaveOptions txtOptions = new TxtSaveOptions
        {
            PreserveTableLayout = true
        };

        // Save the resulting document as a plain‑text file.
        doc.Save("CombinedTable.txt", txtOptions);
    }
}
