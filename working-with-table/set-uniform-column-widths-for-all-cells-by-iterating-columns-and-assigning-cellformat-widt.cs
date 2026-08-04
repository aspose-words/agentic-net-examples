using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder to construct the table.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 3x3 table with sample text.
        Table table = builder.StartTable();

        for (int row = 0; row < 3; row++)
        {
            for (int col = 0; col < 3; col++)
            {
                builder.InsertCell();
                builder.Write($"R{row + 1}C{col + 1}");
            }
            builder.EndRow();
        }

        builder.EndTable();

        // Disable auto‑fit so that explicit widths are respected.
        table.AutoFit(AutoFitBehavior.FixedColumnWidths);

        // Determine the number of columns from the first row.
        int columnCount = table.FirstRow.Cells.Count;
        double uniformWidth = 100.0; // points

        // Iterate each column and set the same width for every cell in that column.
        for (int colIndex = 0; colIndex < columnCount; colIndex++)
        {
            foreach (Row row in table.Rows)
            {
                Cell cell = row.Cells[colIndex];
                cell.CellFormat.Width = uniformWidth;
            }
        }

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "UniformColumnWidths.docx");
        doc.Save(outputPath);

        // Simple validation to ensure the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved correctly.");
    }
}
