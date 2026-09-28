using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Define table dimensions.
        int rows = 5;
        int columns = 4;

        // Build the table.
        builder.StartTable();
        for (int row = 1; row <= rows; row++)
        {
            for (int col = 1; col <= columns; col++)
            {
                builder.InsertCell();
                builder.Writeln($"R{row}C{col}");
                // End the row after the last cell.
                if (col == columns)
                    builder.EndRow();
            }
        }
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        if (table == null)
            throw new InvalidOperationException("Table was not created.");

        // Apply alternating background colors to columns.
        for (int colIndex = 0; colIndex < columns; colIndex++)
        {
            // Choose color based on column index (even = LightGray, odd = White).
            Color bgColor = (colIndex % 2 == 0) ? Color.LightGray : Color.White;

            foreach (Row row in table.Rows)
            {
                Cell cell = row.Cells[colIndex];
                cell.CellFormat.Shading.Texture = TextureIndex.TextureSolid;
                cell.CellFormat.Shading.ForegroundPatternColor = bgColor;
            }
        }

        // Save the document.
        string outputPath = "AlternatingColumnsTable.docx";
        doc.Save(outputPath);

        // Verify that the file was saved.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException($"Failed to create the output file: {outputPath}");
    }
}
