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

        // Build a table with three rows: two with content and one empty.
        Table table = builder.StartTable();

        // Row 1 – contains text.
        builder.InsertCell();
        builder.Write("Row 1, cell 1.");
        builder.EndRow();

        // Row 2 – empty row (no text, default height = 0).
        builder.InsertCell();
        // No text written to the cell.
        builder.EndRow();

        // Row 3 – contains text.
        builder.InsertCell();
        builder.Write("Row 3, cell 1.");
        builder.EndRow();

        builder.EndTable();

        // Save the original document (optional, for verification).
        string originalPath = Path.Combine(Directory.GetCurrentDirectory(), "Original.docx");
        doc.Save(originalPath);

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
        string resultPath = Path.Combine(Directory.GetCurrentDirectory(), "Result.docx");
        doc.Save(resultPath);

        // Simple verification that the output file exists.
        if (!File.Exists(resultPath))
        {
            throw new InvalidOperationException("Result document was not saved correctly.");
        }
    }
}
