using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Tables;

class TableJoinExample
{
    static void Main()
    {
        // Load the source document that already contains two tables.
        Document doc = new Document("InputDocument.docx");

        // Get the first table from the document.
        Table firstTable = doc.FirstSection.Body.Tables[0];

        // Get the second table from the document.
        Table secondTable = (Table)doc.GetChild(NodeType.Table, 1, true);

        // Append all rows from the second table to the first table.
        while (secondTable.HasChildNodes)
            firstTable.Rows.Add(secondTable.FirstRow);

        // Remove the now‑empty second table container.
        secondTable.Remove();

        // Save the resulting document as plain text, preserving the table layout.
        TxtSaveOptions txtOptions = new TxtSaveOptions
        {
            PreserveTableLayout = true
        };
        doc.Save("CombinedTables.txt", txtOptions);
    }
}
