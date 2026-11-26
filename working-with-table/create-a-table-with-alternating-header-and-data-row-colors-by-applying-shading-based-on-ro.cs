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

        // Build the table.
        builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Writeln("Header 1");
        builder.InsertCell();
        builder.Writeln("Header 2");
        builder.EndRow();

        // Add several data rows.
        int dataRowCount = 5;
        for (int i = 0; i < dataRowCount; i++)
        {
            builder.InsertCell();
            builder.Writeln($"Row {i + 1} Col 1");
            builder.InsertCell();
            builder.Writeln($"Row {i + 1} Col 2");
            builder.EndRow();
        }

        // End the table.
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        if (table == null)
            throw new InvalidOperationException("Table was not created.");

        // Apply alternating shading based on row index parity.
        for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++)
        {
            Row row = table.Rows[rowIndex];
            // Even rows get LightGray, odd rows get White.
            Color shade = (rowIndex % 2 == 0) ? Color.LightGray : Color.White;

            // Apply shading to each cell in the row.
            foreach (Cell cell in row.Cells)
            {
                cell.CellFormat.Shading.BackgroundPatternColor = shade;
                // No need to set Texture; default is None.
            }
        }

        // Save the document.
        string outputPath = "TableShading.docx";
        doc.Save(outputPath);

        // Verify that the file was saved.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not created.", outputPath);
    }
}
