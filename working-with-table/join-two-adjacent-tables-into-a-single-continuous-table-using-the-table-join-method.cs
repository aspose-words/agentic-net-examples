using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableJoinExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // ---------- Build the first table ----------
            Table firstTable = builder.StartTable();
            // Row 1
            builder.InsertCell();
            builder.Write("First Table - Row 1, Cell 1");
            builder.InsertCell();
            builder.Write("First Table - Row 1, Cell 2");
            builder.EndRow();
            // Row 2
            builder.InsertCell();
            builder.Write("First Table - Row 2, Cell 1");
            builder.InsertCell();
            builder.Write("First Table - Row 2, Cell 2");
            builder.EndTable(); // firstTable is now complete

            // Insert an empty paragraph to separate the tables (required for adjacency).
            builder.Writeln();

            // ---------- Build the second table ----------
            Table secondTable = builder.StartTable();
            // Row 1
            builder.InsertCell();
            builder.Write("Second Table - Row 1, Cell 1");
            builder.InsertCell();
            builder.Write("Second Table - Row 1, Cell 2");
            builder.EndRow();
            // Row 2
            builder.InsertCell();
            builder.Write("Second Table - Row 2, Cell 1");
            builder.InsertCell();
            builder.Write("Second Table - Row 2, Cell 2");
            builder.EndTable(); // secondTable is now complete

            // Retrieve the table nodes from the document body.
            Table table1 = doc.FirstSection.Body.Tables[0];
            Table table2 = doc.FirstSection.Body.Tables[1];

            // ----- Join the two tables -----
            // Aspose.Words does not provide a Table.Join method. Instead, move all rows
            // from the second table to the first and then remove the empty second table.
            while (table2.HasChildNodes)
            {
                // Append the first row of table2 to table1.
                table1.Rows.Add(table2.FirstRow);
            }
            // Remove the now‑empty second table container.
            table2.Remove();

            // Validation: after joining there should be only one table.
            if (doc.FirstSection.Body.Tables.Count != 1)
                throw new InvalidOperationException("Table join failed: more than one table remains.");

            // Validation: the resulting table should have the combined number of rows (4).
            if (table1.Rows.Count != 4)
                throw new InvalidOperationException($"Table join failed: expected 4 rows, found {table1.Rows.Count}.");

            // Save the resulting document.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "JoinedTables.docx");
            doc.Save(outputPath);
        }
    }
}
