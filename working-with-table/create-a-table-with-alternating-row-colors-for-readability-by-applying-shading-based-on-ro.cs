using System;
using System.Drawing;
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

        // Start building the table.
        builder.StartTable();

        // Create 10 rows with 3 cells each.
        for (int rowIndex = 0; rowIndex < 10; rowIndex++)
        {
            for (int colIndex = 0; colIndex < 3; colIndex++)
            {
                builder.InsertCell();
                builder.Writeln($"Row {rowIndex + 1}, Cell {colIndex + 1}");
            }
            builder.EndRow();
        }

        // End the table construction.
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Apply alternating shading based on row index parity.
        for (int i = 0; i < table.Rows.Count; i++)
        {
            Row row = table.Rows[i];
            // Set shading for each cell in the row because RowFormat does not have a Shading property.
            foreach (Cell cell in row.Cells)
            {
                cell.CellFormat.Shading.Texture = TextureIndex.TextureSolid;
                cell.CellFormat.Shading.BackgroundPatternColor = (i % 2 == 0) ? Color.LightGray : Color.White;
            }
        }

        // Save the document.
        string outputPath = "AlternatingRows.docx";
        doc.Save(outputPath);

        // Validate that the file was saved.
        if (!File.Exists(outputPath))
        {
            throw new Exception("The output document was not created.");
        }
    }
}
