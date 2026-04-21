using System;
using System.IO;
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
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build first sample table.
            Table table1 = builder.StartTable();
            builder.InsertCell();
            builder.Write("Table 1 - Cell 1");
            builder.InsertCell();
            builder.Write("Table 1 - Cell 2");
            builder.EndRow();
            builder.EndTable();

            // Build second sample table.
            Table table2 = builder.StartTable();
            builder.InsertCell();
            builder.Write("Table 2 - Cell 1");
            builder.InsertCell();
            builder.Write("Table 2 - Cell 2");
            builder.EndRow();
            builder.EndTable();

            // Save the sample document.
            string docPath = Path.Combine(Directory.GetCurrentDirectory(), "SampleTables.docx");
            doc.Save(docPath);

            // Retrieve all tables by iterating nodes of type NodeType.Table.
            NodeCollection tableNodes = doc.GetChildNodes(NodeType.Table, true);

            // Output the number of tables found.
            Console.WriteLine($"Total tables found: {tableNodes.Count}");

            // Iterate through each table node.
            for (int i = 0; i < tableNodes.Count; i++)
            {
                Table tbl = (Table)tableNodes[i];
                Console.WriteLine($"Table #{i + 1} has {tbl.Rows.Count} row(s) and {tbl.FirstRow?.Cells.Count ?? 0} column(s).");
            }

            // Optionally write a simple report file.
            string reportPath = Path.Combine(Directory.GetCurrentDirectory(), "TablesReport.txt");
            using (StreamWriter writer = new StreamWriter(reportPath, false))
            {
                writer.WriteLine($"Document: {docPath}");
                writer.WriteLine($"Total tables: {tableNodes.Count}");
                for (int i = 0; i < tableNodes.Count; i++)
                {
                    Table tbl = (Table)tableNodes[i];
                    writer.WriteLine($"Table #{i + 1}: Rows = {tbl.Rows.Count}, Columns = {tbl.FirstRow?.Cells.Count ?? 0}");
                }
            }
        }
    }
}
