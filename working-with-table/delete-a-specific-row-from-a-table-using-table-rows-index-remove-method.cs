using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableRowRemoval
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a 3‑row, 2‑column table.
            Table table = builder.StartTable();

            // Row 1
            builder.InsertCell();
            builder.Write("Row 1, Cell 1.");
            builder.InsertCell();
            builder.Write("Row 1, Cell 2.");
            builder.EndRow();

            // Row 2
            builder.InsertCell();
            builder.Write("Row 2, Cell 1.");
            builder.InsertCell();
            builder.Write("Row 2, Cell 2.");
            builder.EndRow();

            // Row 3
            builder.InsertCell();
            builder.Write("Row 3, Cell 1.");
            builder.InsertCell();
            builder.Write("Row 3, Cell 2.");
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Delete the second row (index 1) using the Rows collection.
            if (table.Rows.Count > 1)
            {
                table.Rows[1].Remove();
            }
            else
            {
                throw new InvalidOperationException("The table does not contain enough rows to delete the specified index.");
            }

            // Verify that the table now has exactly two rows.
            if (table.Rows.Count != 2)
                throw new Exception("Row removal failed: unexpected row count.");

            // Save the resulting document.
            const string outputPath = "DeletedRowTable.docx";
            doc.Save(outputPath);

            // Inform that the operation completed successfully.
            Console.WriteLine($"Document saved to '{outputPath}'. Remaining rows: {table.Rows.Count}");
        }
    }
}
