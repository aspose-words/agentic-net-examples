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

            // Build first table (2 rows x 2 columns).
            Table table1 = builder.StartTable();
            builder.InsertCell();
            builder.Write("R1C1");
            builder.InsertCell();
            builder.Write("R1C2");
            builder.EndRow();

            builder.InsertCell();
            builder.Write("R2C1");
            builder.InsertCell();
            builder.Write("R2C2");
            builder.EndRow();
            builder.EndTable();

            // Build second table (1 row x 3 columns).
            Table table2 = builder.StartTable();
            builder.InsertCell();
            builder.Write("A");
            builder.InsertCell();
            builder.Write("B");
            builder.InsertCell();
            builder.Write("C");
            builder.EndRow();
            builder.EndTable();

            // Save the document to a local file (required by the rules).
            string filePath = Path.Combine(Directory.GetCurrentDirectory(), "Sample.docx");
            doc.Save(filePath);

            // Load the document back (demonstrates load workflow).
            Document loadedDoc = new Document(filePath);

            // Retrieve all tables by iterating nodes of type NodeType.Table.
            NodeCollection tableNodes = loadedDoc.GetChildNodes(NodeType.Table, true);

            // Iterate through the tables and output basic information.
            for (int i = 0; i < tableNodes.Count; i++)
            {
                Table tbl = (Table)tableNodes[i];
                int rowCount = tbl.Rows.Count;
                int firstRowCellCount = tbl.FirstRow?.Cells.Count ?? 0;

                Console.WriteLine($"Table {i}: Rows = {rowCount}, Cells in first row = {firstRowCellCount}");
            }

            // The program finishes without waiting for user input.
        }
    }
}
