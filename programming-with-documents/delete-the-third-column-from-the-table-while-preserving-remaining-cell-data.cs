using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a table with 3 columns and 3 rows.
        Table table = builder.StartTable();

        // Header row
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.InsertCell();
        builder.Write("Header 3");
        builder.EndRow();

        // First data row
        builder.InsertCell();
        builder.Write("R1C1");
        builder.InsertCell();
        builder.Write("R1C2");
        builder.InsertCell();
        builder.Write("R1C3");
        builder.EndRow();

        // Second data row
        builder.InsertCell();
        builder.Write("R2C1");
        builder.InsertCell();
        builder.Write("R2C2");
        builder.InsertCell();
        builder.Write("R2C3");
        builder.EndRow();

        builder.EndTable();

        // Delete the third column (zero‑based index 2) while preserving other cell data.
        int columnIndexToRemove = 2;
        foreach (Row row in table.Rows)
        {
            // Ensure the row has enough cells before attempting removal.
            if (row.Cells.Count > columnIndexToRemove)
                row.Cells.RemoveAt(columnIndexToRemove);
        }

        // Save the modified document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);
    }
}
