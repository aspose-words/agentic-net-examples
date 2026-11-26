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

        // Add a heading.
        builder.Writeln("Table with alternating column shading:");

        // Define table dimensions.
        int rows = 5;
        int columns = 4;

        // Build the table.
        builder.StartTable();
        for (int r = 0; r < rows; r++)
        {
            for (int c = 0; c < columns; c++)
            {
                builder.InsertCell();
                builder.Writeln($"R{r + 1}C{c + 1}");
            }
            builder.EndRow();
        }
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        if (table == null)
            throw new InvalidOperationException("Table was not created.");

        // Define two alternating colors.
        Color colorEven = Color.LightGray;
        Color colorOdd = Color.White;

        // Apply shading to each column alternately.
        for (int colIdx = 0; colIdx < columns; colIdx++)
        {
            Color shade = (colIdx % 2 == 0) ? colorEven : colorOdd;
            foreach (Row row in table.Rows)
            {
                Cell cell = row.Cells[colIdx];
                cell.CellFormat.Shading.ForegroundPatternColor = shade;
                cell.CellFormat.Shading.Texture = TextureIndex.TextureSolid;
            }
        }

        // Save the document.
        string outputPath = "AlternatingColumnsTable.docx";
        doc.Save(outputPath);

        // Verify that the file was saved.
        if (!File.Exists(outputPath))
            throw new Exception("The output file was not created.");

        // Optionally, inform that the process completed.
        Console.WriteLine($"Document saved to '{Path.GetFullPath(outputPath)}'.");
    }
}
