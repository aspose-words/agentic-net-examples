using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableJoinExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // ---------- First table ----------
            Table firstTable = builder.StartTable();
            builder.InsertCell();
            builder.Write("First Table - Row 1, Cell 1");
            builder.InsertCell();
            builder.Write("First Table - Row 1, Cell 2");
            builder.EndRow();

            builder.InsertCell();
            builder.Write("First Table - Row 2, Cell 1");
            builder.InsertCell();
            builder.Write("First Table - Row 2, Cell 2");
            builder.EndRow();

            builder.EndTable(); // Cursor is now positioned after the first table.

            // ---------- Second table (adjacent) ----------
            Table secondTable = builder.StartTable();
            builder.InsertCell();
            builder.Write("Second Table - Row 1, Cell 1");
            builder.InsertCell();
            builder.Write("Second Table - Row 1, Cell 2");
            builder.EndRow();

            builder.InsertCell();
            builder.Write("Second Table - Row 2, Cell 1");
            builder.InsertCell();
            builder.Write("Second Table - Row 2, Cell 2");
            builder.EndRow();

            builder.EndTable(); // Cursor is now positioned after the second table.

            // Retrieve the two tables from the document body.
            Table table1 = doc.FirstSection.Body.Tables[0];
            Table table2 = doc.FirstSection.Body.Tables[1];

            // Join the second table into the first one by moving all rows.
            while (table2.HasChildNodes)
                table1.Rows.Add(table2.FirstRow);
            // Remove the now‑empty second table container.
            table2.Remove();

            // Verify that only one table remains.
            if (doc.FirstSection.Body.Tables.Count != 1)
                throw new InvalidOperationException("Tables were not joined correctly.");

            // Save the resulting document.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "JoinedTables.docx");
            doc.Save(outputPath);

            // Inform the user (no interactive input required).
            Console.WriteLine($"Document saved to: {outputPath}");
        }
    }
}
