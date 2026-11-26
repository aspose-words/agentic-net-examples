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

        // Build a table with 5 rows and 2 columns.
        builder.StartTable();
        for (int row = 1; row <= 5; row++)
        {
            for (int col = 1; col <= 2; col++)
            {
                builder.InsertCell();
                builder.Writeln($"Row {row}, Cell {col}");
            }
            builder.EndRow();
        }
        builder.EndTable();

        // Retrieve the first (and only) table in the document.
        Table originalTable = (Table)doc.GetChild(NodeType.Table, 0, true);
        if (originalTable == null)
            throw new InvalidOperationException("Table was not created.");

        // Index of the row after which we split (zero‑based).
        // Split at the third row => rows 0‑2 stay in the original table.
        int splitRowIndex = 2;

        // Create a new empty table that copies the formatting of the original.
        Table newTable = (Table)originalTable.Clone(false); // shallow clone, no rows

        // Move rows that belong to the second part into the new table.
        while (originalTable.Rows.Count > splitRowIndex + 1)
        {
            // Row immediately after the split index.
            Row movingRow = originalTable.Rows[splitRowIndex + 1];
            // Remove it from the original table.
            originalTable.Rows.RemoveAt(splitRowIndex + 1);
            // Add it to the new table.
            newTable.Rows.Add(movingRow);
        }

        // Insert the new table into the document right after the original table.
        originalTable.ParentNode.InsertAfter(newTable, originalTable);

        // Insert an empty paragraph between the two tables for clarity.
        Paragraph separator = new Paragraph(doc);
        originalTable.ParentNode.InsertAfter(separator, originalTable);

        // Save the document.
        string outputPath = "SplitTable.docx";
        doc.Save(outputPath);

        // Verify that the file was saved.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not created.", outputPath);
    }
}
