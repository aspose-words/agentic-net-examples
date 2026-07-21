using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableSplitExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a table with 10 rows and 3 columns.
            Table table = builder.StartTable();

            for (int row = 0; row < 10; row++)
            {
                for (int col = 0; col < 3; col++)
                {
                    builder.InsertCell();
                    builder.Write($"Row {row + 1}, Col {col + 1}");
                }
                builder.EndRow();
            }

            // Finish the table construction.
            builder.EndTable();

            // Split the table after the 5th row (zero‑based index 4).
            int splitRowIndex = 4;

            // Create a new empty table that will hold the rows after the split point.
            Table newTable = new Table(doc);

            // Move rows that come after the split index from the original table to the new table.
            while (table.Rows.Count > splitRowIndex + 1)
            {
                // The row to move is the one immediately after the split index.
                Row rowToMove = table.Rows[splitRowIndex + 1];
                // Detach the row from the original table.
                rowToMove.Remove();
                // Append the row to the new table.
                newTable.Rows.Add(rowToMove);
            }

            // Insert a paragraph between the two tables to make the split visible.
            builder.Writeln();
            builder.Writeln("Table was split here.");

            // Insert the new table after the paragraph we have just added.
            // The builder's cursor is positioned inside the paragraph we just wrote.
            // InsertAfter on the parent (the body) places the new table correctly.
            Node paragraphNode = builder.CurrentParagraph;
            paragraphNode.ParentNode.InsertAfter(newTable, paragraphNode);

            // Save the resulting document.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "SplitTable.docx");
            doc.Save(outputPath);

            // Output verification information.
            Console.WriteLine($"Document saved to: {outputPath}");
            Console.WriteLine($"Original table rows after split: {table.Rows.Count}");
            Console.WriteLine($"New table rows after split: {newTable.Rows.Count}");
        }
    }
}
