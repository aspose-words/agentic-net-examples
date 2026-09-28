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

        int rowCount = 6; // Total number of rows (must be even for pairing).

        // Begin a table with a single column.
        builder.StartTable();

        for (int i = 0; i < rowCount; i++)
        {
            // Insert a cell for the current row.
            builder.InsertCell();

            // Add some text to the cell.
            builder.Writeln($"Row {i + 1}");

            // Retrieve the cell that was just created.
            Cell cell = (Cell)builder.CurrentParagraph.ParentNode;

            // Merge every second cell vertically:
            // - First cell of each pair starts the merge.
            // - Second cell of each pair continues the merge.
            if (i % 2 == 0)
                cell.CellFormat.VerticalMerge = CellMerge.First;
            else
                cell.CellFormat.VerticalMerge = CellMerge.Previous;

            // End the current row.
            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Save the document to disk.
        string fileName = "TableMerge.docx";
        doc.Save(fileName);

        // Verify that the file was created.
        if (!File.Exists(fileName))
            throw new Exception("Failed to create the output document.");
    }
}
