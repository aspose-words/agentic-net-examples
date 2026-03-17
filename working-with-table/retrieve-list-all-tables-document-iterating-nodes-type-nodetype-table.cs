using Aspose.Words;
using Aspose.Words.Tables;
using System;

class Program
{
    static void Main()
    {
        // Load the document (adjust the path as needed)
        Document doc = new Document("Input.docx");

        // Retrieve all Table nodes in the document (deep traversal)
        NodeCollection tableNodes = doc.GetChildNodes(NodeType.Table, true);

        // Iterate over the collection and work with each Table
        for (int i = 0; i < tableNodes.Count; i++)
        {
            Table table = (Table)tableNodes[i];
            int rowCount = table.Rows.Count;
            int columnCount = table.FirstRow?.Count ?? 0;
            Console.WriteLine($"Table #{i}: {rowCount} rows, {columnCount} columns");
        }

        // Save the document (no modifications made in this example)
        doc.Save("Output.docx");
    }
}
