using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Build the first sample table.
            DocumentBuilder builder = new DocumentBuilder(doc);
            Table table1 = builder.StartTable();
            builder.InsertCell();
            builder.Write("Table 1, Cell 1");
            builder.InsertCell();
            builder.Write("Table 1, Cell 2");
            builder.EndRow();
            builder.EndTable();

            // Build the second sample table.
            Table table2 = builder.StartTable();
            builder.InsertCell();
            builder.Write("Table 2, Cell 1");
            builder.InsertCell();
            builder.Write("Table 2, Cell 2");
            builder.EndRow();
            builder.EndTable();

            // Save the document to the local file system.
            const string outputPath = "Output.docx";
            doc.Save(outputPath);

            // Load the document back (demonstrates load rule usage).
            Document loadedDoc = new Document(outputPath);

            // Retrieve all tables by iterating nodes of type NodeType.Table.
            NodeCollection tableNodes = loadedDoc.GetChildNodes(NodeType.Table, true);
            Console.WriteLine($"Total tables found: {tableNodes.Count}");

            for (int i = 0; i < tableNodes.Count; i++)
            {
                Table tbl = (Table)tableNodes[i];
                Console.WriteLine($"Table #{i + 1} has {tbl.Rows.Count} row(s) and {tbl.FirstRow?.Cells.Count ?? 0} column(s).");
            }
        }
    }
}
