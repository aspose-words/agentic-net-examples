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

        // Begin the table.
        builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Writeln("Header 1");
        builder.InsertCell();
        builder.Writeln("Header 2");
        builder.EndRow();

        // Add several data rows.
        int dataRowCount = 6;
        for (int i = 0; i < dataRowCount; i++)
        {
            builder.InsertCell();
            builder.Writeln($"Row {i + 1} - Col 1");
            builder.InsertCell();
            builder.Writeln($"Row {i + 1} - Col 2");
            builder.EndRow();
        }

        // End the table.
        builder.EndTable();

        // Retrieve the created table from the document.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Apply shading to the header row cells.
        Row headerRow = table.Rows[0];
        foreach (Cell cell in headerRow.Cells)
        {
            // Set solid gray background.
            cell.CellFormat.Shading.Texture = TextureIndex.TextureSolid;
            cell.CellFormat.Shading.BackgroundPatternColor = Color.Gray;
        }

        // Apply alternating shading to data rows based on row index parity.
        for (int i = 1; i < table.Rows.Count; i++)
        {
            Row row = table.Rows[i];
            bool isEvenDataRow = ((i - 1) % 2 == 0); // (i-1) because first data row is index 1.

            foreach (Cell cell in row.Cells)
            {
                if (isEvenDataRow)
                {
                    // Light gray background for even rows.
                    cell.CellFormat.Shading.Texture = TextureIndex.TextureSolid;
                    cell.CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
                }
                else
                {
                    // Clear shading for odd rows.
                    cell.CellFormat.Shading.Texture = TextureIndex.TextureNone;
                    cell.CellFormat.Shading.BackgroundPatternColor = Color.Empty;
                }
            }
        }

        // Save the document to a file.
        string outputPath = "AlternatingRows.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }
    }
}
