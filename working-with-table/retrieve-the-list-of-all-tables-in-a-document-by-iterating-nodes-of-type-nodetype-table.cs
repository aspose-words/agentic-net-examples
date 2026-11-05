using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace WorkingWithTableExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build the first table (2 rows x 2 columns).
            builder.StartTable();
            builder.InsertCell();
            builder.Writeln("Table1 Row1 Cell1");
            builder.InsertCell();
            builder.Writeln("Table1 Row1 Cell2");
            builder.EndRow();

            builder.InsertCell();
            builder.Writeln("Table1 Row2 Cell1");
            builder.InsertCell();
            builder.Writeln("Table1 Row2 Cell2");
            builder.EndRow();
            builder.EndTable();

            // Add a paragraph between tables.
            builder.Writeln();

            // Build the second table (1 row x 3 columns).
            builder.StartTable();
            builder.InsertCell();
            builder.Writeln("Table2 Cell1");
            builder.InsertCell();
            builder.Writeln("Table2 Cell2");
            builder.InsertCell();
            builder.Writeln("Table2 Cell3");
            builder.EndRow();
            builder.EndTable();

            // Save the document to a local file.
            string outputPath = "SampleTables.docx";
            doc.Save(outputPath);

            // Verify that the file was saved.
            if (!File.Exists(outputPath))
                throw new Exception("Failed to save the document.");

            // Load the document (demonstrates loading workflow).
            Document loadedDoc = new Document(outputPath);

            // Retrieve all tables in the document.
            NodeCollection tableNodes = loadedDoc.GetChildNodes(NodeType.Table, true);

            Console.WriteLine($"Total tables found: {tableNodes.Count}");

            // Iterate through each table and output its details.
            for (int i = 0; i < tableNodes.Count; i++)
            {
                Table table = (Table)tableNodes[i];
                Console.WriteLine($"Table {i + 1}: Rows = {table.Rows.Count}, Columns = {GetColumnCount(table)}");
            }
        }

        // Helper method to determine the maximum column count in a table.
        private static int GetColumnCount(Table table)
        {
            int max = 0;
            foreach (Row row in table.Rows)
            {
                int cells = row.Cells.Count;
                if (cells > max)
                    max = cells;
            }
            return max;
        }
    }
}
