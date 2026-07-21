using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableColumnRemoval
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a 3x3 table.
            Table table = builder.StartTable();

            // First row.
            builder.InsertCell();
            builder.Write("R1C1");
            builder.InsertCell();
            builder.Write("R1C2");
            builder.InsertCell();
            builder.Write("R1C3");
            builder.EndRow();

            // Second row.
            builder.InsertCell();
            builder.Write("R2C1");
            builder.InsertCell();
            builder.Write("R2C2");
            builder.InsertCell();
            builder.Write("R2C3");
            builder.EndRow();

            // Third row.
            builder.InsertCell();
            builder.Write("R3C1");
            builder.InsertCell();
            builder.Write("R3C2");
            builder.InsertCell();
            builder.Write("R3C3");
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Delete the second column (index 1) by removing the cell at that index from each row.
            int columnIndexToRemove = 1;
            if (table.Rows.Count > 0 && table.Rows[0].Cells.Count > columnIndexToRemove)
            {
                foreach (Row row in table.Rows)
                {
                    // Ensure the current row has enough cells before removal.
                    if (row.Cells.Count > columnIndexToRemove)
                        row.Cells.RemoveAt(columnIndexToRemove);
                }
            }

            // Save the resulting document.
            doc.Save("DeleteColumn.docx");
        }
    }
}
