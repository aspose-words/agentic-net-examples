using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class TableShadingExample
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table.
        builder.StartTable();

        // Build a 3x4 table (3 rows, 4 columns) with sample text.
        int rows = 3;
        int columns = 4;
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

        // Retrieve the created table.
        Table table = doc.FirstSection.Body.Tables[0];

        // Iterate through each cell and apply background color based on column index.
        for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++)
        {
            Row row = table.Rows[rowIndex];
            for (int cellIndex = 0; cellIndex < row.Cells.Count; cellIndex++)
            {
                Cell cell = row.Cells[cellIndex];

                // Choose a color: even columns LightGray, odd columns LightBlue.
                Color bgColor = (cellIndex % 2 == 0) ? Color.LightGray : Color.LightBlue;

                // Apply solid shading with the chosen background color.
                // The default texture is None, so we only need to set the background color.
                cell.CellFormat.Shading.BackgroundPatternColor = bgColor;
            }
        }

        // Save the document.
        string outputPath = "TableShading.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }
    }
}
