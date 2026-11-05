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

        // Build a simple 3‑column, 2‑row table.
        builder.StartTable();

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
        builder.EndRow();

        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Desired uniform column width (points).
        double uniformWidth = 100.0;

        // Determine the number of columns from the first row.
        int columnCount = table.Rows[0].Cells.Count;

        // Iterate each column and set the width for every cell in that column.
        for (int col = 0; col < columnCount; col++)
        {
            foreach (Row row in table.Rows)
            {
                Cell cell = row.Cells[col];
                cell.CellFormat.Width = uniformWidth;
            }
        }

        // Save the document.
        string outputPath = "UniformColumnWidths.docx";
        doc.Save(outputPath);

        // Verify that the file was saved.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Failed to save the output document.");

        // Indicate successful completion.
        Console.WriteLine($"Document saved to: {Path.GetFullPath(outputPath)}");
    }
}
