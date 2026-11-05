using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table with a single column.
        builder.StartTable();

        int totalRows = 8; // Number of rows in the table.

        // Insert rows with a single cell each.
        for (int i = 0; i < totalRows; i++)
        {
            builder.InsertCell();
            builder.Writeln($"Row {i + 1}");
            builder.EndRow();
        }

        // End the table.
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChildNodes(NodeType.Table, true)[0];

        // Merge every second cell vertically (rows 2&3, 4&5, etc.).
        // Loop stops before the last row to avoid out‑of‑range access.
        for (int rowIndex = 1; rowIndex < totalRows - 1; rowIndex += 2)
        {
            Cell firstCell = table.Rows[rowIndex].Cells[0];
            Cell secondCell = table.Rows[rowIndex + 1].Cells[0];

            // Mark the first cell as the start of a vertical merge.
            firstCell.CellFormat.VerticalMerge = CellMerge.First;
            // Mark the second cell to merge with the previous cell.
            secondCell.CellFormat.VerticalMerge = CellMerge.Previous;
        }

        // Save the document.
        string outputPath = "TableMerge.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");

        // Indicate successful completion.
        Console.WriteLine($"Document saved to {Path.GetFullPath(outputPath)}");
    }
}
