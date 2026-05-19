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

        // Build a table with three rows.
        // The second row will be given a height of zero to simulate an empty row.
        Table table = builder.StartTable();

        // ----- Row 1 -----
        builder.InsertCell();
        builder.Write("Row 1, Cell 1");
        builder.InsertCell();
        builder.Write("Row 1, Cell 2");
        builder.EndRow();

        // ----- Row 2 (empty, height = 0) -----
        builder.RowFormat.Height = 0;
        builder.RowFormat.HeightRule = HeightRule.Exactly; // enforce the exact height
        builder.InsertCell();
        builder.Write(string.Empty);
        builder.InsertCell();
        builder.Write(string.Empty);
        builder.EndRow();

        // ----- Row 3 -----
        // Reset the row formatting so the next rows are normal.
        builder.RowFormat.Height = 0;
        builder.RowFormat.HeightRule = HeightRule.Auto;
        builder.InsertCell();
        builder.Write("Row 3, Cell 1");
        builder.InsertCell();
        builder.Write("Row 3, Cell 2");
        builder.EndRow();

        builder.EndTable();

        // Remove all rows whose height is zero.
        // Iterate backwards to avoid index shifting when rows are removed.
        for (int i = table.Rows.Count - 1; i >= 0; i--)
        {
            Row row = table.Rows[i];
            if (row.RowFormat.Height == 0)
            {
                row.Remove();
            }
        }

        // Save the resulting document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Result.docx");
        doc.Save(outputPath);
    }
}
