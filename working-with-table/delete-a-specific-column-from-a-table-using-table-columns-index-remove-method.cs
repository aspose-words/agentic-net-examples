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

        // Build a sample table with 3 columns and 3 rows.
        builder.StartTable();

        // Header row
        builder.InsertCell();
        builder.Writeln("Header 1");
        builder.InsertCell();
        builder.Writeln("Header 2");
        builder.InsertCell();
        builder.Writeln("Header 3");
        builder.EndRow();

        // First data row
        builder.InsertCell();
        builder.Writeln("R1C1");
        builder.InsertCell();
        builder.Writeln("R1C2");
        builder.InsertCell();
        builder.Writeln("R1C3");
        builder.EndRow();

        // Second data row
        builder.InsertCell();
        builder.Writeln("R2C1");
        builder.InsertCell();
        builder.Writeln("R2C2");
        builder.InsertCell();
        builder.Writeln("R2C3");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Retrieve the first table in the document.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Delete the second column (index 1) by removing the cell at that index from each row.
        int columnIndexToRemove = 1; // zero‑based index for the second column

        // Ensure the table has enough columns to remove.
        if (table.Rows.Count > 0 && table.Rows[0].Cells.Count > columnIndexToRemove)
        {
            foreach (Row row in table.Rows)
            {
                // Guard against rows that might have fewer cells (unlikely in a regular table).
                if (row.Cells.Count > columnIndexToRemove)
                {
                    row.Cells.RemoveAt(columnIndexToRemove);
                }
            }
        }
        else
        {
            throw new InvalidOperationException("The table does not have enough columns to remove.");
        }

        // Save the modified document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new FileNotFoundException("The output document was not saved correctly.", outputPath);
        }
    }
}
