using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main(string[] args)
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new table.
        builder.StartTable();

        // Insert three cells with sample text.
        builder.InsertCell();
        builder.Write("Column 1");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("Column 2");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("Column 3");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Disable AutoFit to use fixed layout.
        table.AutoFit(AutoFitBehavior.FixedColumnWidths);

        // Set explicit widths for each column (in points).
        // Widths are applied to the cells in the first row.
        double[] columnWidths = { 100.0, 150.0, 200.0 };
        Row firstRow = table.FirstRow;

        for (int i = 0; i < firstRow.Cells.Count && i < columnWidths.Length; i++)
        {
            Cell cell = firstRow.Cells[i];
            cell.CellFormat.Width = columnWidths[i];
        }

        // Save the document.
        string outputPath = "FixedLayoutTable.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }
    }
}
