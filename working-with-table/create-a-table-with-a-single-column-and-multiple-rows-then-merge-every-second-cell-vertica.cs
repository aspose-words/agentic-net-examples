using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table that will have a single column.
        Table table = builder.StartTable();

        // Define the number of rows to create.
        int totalRows = 7; // Example: 7 rows (will create 3 merged groups and 1 single row)

        for (int i = 0; i < totalRows; i++)
        {
            // Insert a cell for the current row.
            builder.InsertCell();

            // Merge every second cell vertically.
            if (i % 2 == 0) // First cell of a pair – start of a vertical merge.
            {
                builder.CellFormat.VerticalMerge = CellMerge.First;
                builder.Write($"Group {(i / 2) + 1}");
            }
            else // Second cell of a pair – continues the previous merge.
            {
                builder.CellFormat.VerticalMerge = CellMerge.Previous;
                // No text needed for the merged‑into cell.
            }

            // End the current row.
            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Save the document to a file.
        string outputPath = "MergedTable.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The document was not saved correctly.");
    }
}
