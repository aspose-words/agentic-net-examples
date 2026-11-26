using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a 3x3 table.
        builder.StartTable();
        for (int row = 0; row < 3; row++)
        {
            for (int col = 0; col < 3; col++)
            {
                builder.InsertCell();
                builder.Writeln($"R{row + 1}C{col + 1}");
            }
            builder.EndRow();
        }
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Apply background colors based on column index.
        for (int rowIdx = 0; rowIdx < table.Rows.Count; rowIdx++)
        {
            Row row = table.Rows[rowIdx];
            for (int colIdx = 0; colIdx < row.Cells.Count; colIdx++)
            {
                Cell cell = row.Cells[colIdx];
                Color bgColor;
                switch (colIdx)
                {
                    case 0:
                        bgColor = Color.LightCoral;
                        break;
                    case 1:
                        bgColor = Color.LightGreen;
                        break;
                    case 2:
                        bgColor = Color.LightBlue;
                        break;
                    default:
                        bgColor = Color.White;
                        break;
                }

                // Set solid shading for the cell.
                // No need to set Texture; default is solid (None).
                cell.CellFormat.Shading.ForegroundPatternColor = bgColor; // Desired background color
                cell.CellFormat.Shading.BackgroundPatternColor = Color.Empty; // No background pattern color
            }
        }

        // Save the document.
        string outputPath = "TableWithShading.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");
    }
}
