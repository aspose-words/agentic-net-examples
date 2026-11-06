using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace DeleteTableColumnExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a sample table with 3 columns and 3 rows.
            builder.StartTable();

            // Header row.
            builder.InsertCell();
            builder.Writeln("Header 1");
            builder.InsertCell();
            builder.Writeln("Header 2");
            builder.InsertCell();
            builder.Writeln("Header 3");
            builder.EndRow();

            // First data row.
            builder.InsertCell();
            builder.Writeln("R1C1");
            builder.InsertCell();
            builder.Writeln("R1C2");
            builder.InsertCell();
            builder.Writeln("R1C3");
            builder.EndRow();

            // Second data row.
            builder.InsertCell();
            builder.Writeln("R2C1");
            builder.InsertCell();
            builder.Writeln("R2C2");
            builder.InsertCell();
            builder.Writeln("R2C3");
            builder.EndRow();

            builder.EndTable();

            // Locate the first table in the document.
            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

            // Delete the third column (zero‑based index 2) by removing the cell
            // at that index from each row.
            int columnIndexToRemove = 2;
            foreach (Row row in table.Rows)
            {
                // Ensure the row actually has the column we want to delete.
                if (row.Cells.Count > columnIndexToRemove)
                {
                    // Remove the cell; this also removes its content.
                    row.Cells[columnIndexToRemove].Remove();
                }
            }

            // Save the modified document.
            string outputPath = "Output.docx";
            doc.Save(outputPath);
        }
    }
}
