using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableTraversal
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build first sample table.
            builder.StartTable();
            builder.InsertCell();
            builder.Write("Table 1 - Cell 1");
            builder.InsertCell();
            builder.Write("Table 1 - Cell 2");
            builder.EndRow();
            builder.EndTable();

            // Add a paragraph between tables.
            builder.Writeln();

            // Build second sample table.
            builder.StartTable();
            builder.InsertCell();
            builder.Write("Table 2 - Cell 1");
            builder.InsertCell();
            builder.Write("Table 2 - Cell 2");
            builder.InsertCell();
            builder.Write("Table 2 - Cell 3");
            builder.EndRow();
            builder.EndTable();

            // Save the document to a local file.
            string filePath = "Sample.docx";
            doc.Save(filePath);

            // Load the document back (demonstrates load workflow).
            Document loadedDoc = new Document(filePath);

            // Retrieve all tables by iterating nodes of type NodeType.Table.
            NodeCollection tableNodes = loadedDoc.GetChildNodes(NodeType.Table, true);

            // Output information about each table.
            Console.WriteLine($"Total tables found: {tableNodes.Count}");
            for (int i = 0; i < tableNodes.Count; i++)
            {
                Table table = (Table)tableNodes[i];
                int rowCount = table.Rows.Count;
                int cellCount = 0;
                foreach (Row row in table.Rows)
                {
                    cellCount += row.Cells.Count;
                }

                Console.WriteLine($"Table {i + 1}: Rows = {rowCount}, Cells = {cellCount}");
            }

            // Optionally write a simple report file.
            string reportPath = "TablesReport.txt";
            using (StreamWriter writer = new StreamWriter(reportPath))
            {
                writer.WriteLine($"Document: {Path.GetFullPath(filePath)}");
                writer.WriteLine($"Total tables: {tableNodes.Count}");
                for (int i = 0; i < tableNodes.Count; i++)
                {
                    Table table = (Table)tableNodes[i];
                    writer.WriteLine($"Table {i + 1}: Rows = {table.Rows.Count}, Cells = {GetCellCount(table)}");
                }
            }

            // Verify that the report file was created.
            if (!File.Exists(reportPath))
                throw new InvalidOperationException("Report file was not created.");

            // Helper local function to count cells in a table.
            int GetCellCount(Table tbl)
            {
                int count = 0;
                foreach (Row r in tbl.Rows)
                    count += r.Cells.Count;
                return count;
            }
        }
    }
}
