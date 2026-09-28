using System;
using System.Drawing;
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

        // Start a table.
        builder.StartTable();

        int rowCount = 10;
        int columnCount = 3;

        // Build the table rows and cells.
        for (int i = 0; i < rowCount; i++)
        {
            for (int j = 0; j < columnCount; j++)
            {
                builder.InsertCell();
                builder.Writeln($"Row {i + 1}, Cell {j + 1}");
            }
            builder.EndRow();
        }

        // End the table.
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChildNodes(NodeType.Table, true)[0];

        // Apply shading to every second row (index 1,3,5,...).
        for (int i = 0; i < table.Rows.Count; i++)
        {
            if (i % 2 == 1) // every second row (1‑based even rows)
            {
                Row row = table.Rows[i];
                foreach (Cell cell in row.Cells)
                {
                    cell.CellFormat.Shading.Texture = TextureIndex.TextureSolid;
                    cell.CellFormat.Shading.ForegroundPatternColor = Color.LightGray;
                }
            }
        }

        // Save the document.
        string outputPath = "AlternatingRows.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not saved correctly.");
    }
}
