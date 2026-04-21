using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a table with 3 columns and 2 rows.
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
        builder.EndTable();

        // Delete the second column (index 1) by removing the cell at that index from each row.
        int columnIndexToRemove = 1;
        if (table.Rows.Count > 0 && columnIndexToRemove >= 0 && columnIndexToRemove < table.Rows[0].Cells.Count)
        {
            foreach (Row row in table.Rows)
            {
                // Ensure the row has enough cells before attempting removal.
                if (columnIndexToRemove < row.Cells.Count)
                {
                    row.Cells.RemoveAt(columnIndexToRemove);
                }
            }
        }

        // Save the resulting document.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "DeletedColumn.docx");
        doc.Save(outputPath);
    }
}
