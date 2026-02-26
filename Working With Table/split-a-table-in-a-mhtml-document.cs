using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

class SplitTableMhtml
{
    static void Main()
    {
        // Load the source MHTML document.
        Document srcDoc = new Document("Input.mhtml");

        // Retrieve all tables in the document.
        NodeCollection tableNodes = srcDoc.GetChildNodes(NodeType.Table, true);

        // Create an output folder for the split tables.
        string outFolder = "SplitTables";
        Directory.CreateDirectory(outFolder);

        // Iterate over each table, clone it, and save it as a separate MHTML file.
        for (int i = 0; i < tableNodes.Count; i++)
        {
            // Clone the table (deep clone to include all child nodes).
            Table clonedTable = (Table)tableNodes[i].Clone(true);

            // Create a new empty document.
            Document newDoc = new Document();

            // Append the cloned table to the new document's body.
            newDoc.FirstSection.Body.AppendChild(clonedTable);

            // Save the new document as MHTML.
            string outPath = Path.Combine(outFolder, $"Table_{i + 1}.mhtml");
            newDoc.Save(outPath, SaveFormat.Mhtml);
        }
    }
}
