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

        // Start a table.
        builder.StartTable();

        // Define number of rows and columns.
        int rows = 4;
        int columns = 3;

        // Populate the table with sample text.
        for (int r = 0; r < rows; r++)
        {
            for (int c = 0; c < columns; c++)
            {
                builder.InsertCell();
                builder.Writeln($"R{r + 1}C{c + 1}");
            }
            builder.EndRow();
        }

        // End the table.
        builder.EndTable();

        // Apply background colors based on column index.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        for (int r = 0; r < table.Rows.Count; r++)
        {
            Row row = table.Rows[r];
            for (int c = 0; c < row.Cells.Count; c++)
            {
                Cell cell = row.Cells[c];
                // Choose a color per column.
                Color bgColor = c switch
                {
                    0 => Color.LightCoral,   // First column
                    1 => Color.LightGreen,   // Second column
                    2 => Color.LightBlue,    // Third column
                    _ => Color.White
                };
                // Apply solid shading.
                cell.CellFormat.Shading.BackgroundPatternColor = bgColor;
                cell.CellFormat.Shading.Texture = TextureIndex.TextureSolid;
            }
        }

        // Save the document.
        string outputPath = "TableWithColumnShading.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");

        // Optional: reopen and ensure the table still exists (basic validation).
        Document loadedDoc = new Document(outputPath);
        Table loadedTable = (Table)loadedDoc.GetChild(NodeType.Table, 0, true);
        if (loadedTable == null || loadedTable.Rows.Count != rows)
            throw new InvalidOperationException("The saved document does not contain the expected table.");

        // Program ends without waiting for user input.
    }
}
