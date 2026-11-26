using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a heading.
        builder.Writeln("Table with alternating column shading:");

        // Define table dimensions.
        int rows = 4;
        int columns = 5;

        // Build the table.
        builder.StartTable();
        for (int r = 1; r <= rows; r++)
        {
            for (int c = 1; c <= columns; c++)
            {
                builder.InsertCell();
                builder.Writeln($"R{r}C{c}");
            }
            builder.EndRow();
        }
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChildNodes(NodeType.Table, true)[0];

        // Apply shading to every second column (index 1, 3, ...).
        for (int colIndex = 0; colIndex < columns; colIndex++)
        {
            if (colIndex % 2 == 1) // every second column
            {
                foreach (Row row in table.Rows)
                {
                    Cell cell = row.Cells[colIndex];
                    cell.CellFormat.Shading.Texture = TextureIndex.TextureSolid;
                    cell.CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
                }
            }
        }

        // Save the document.
        string outputPath = "AlternatingColumnShading.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not saved correctly.");
    }
}
