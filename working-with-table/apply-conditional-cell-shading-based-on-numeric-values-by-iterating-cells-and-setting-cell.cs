using System;
using System.IO;
using System.Drawing;
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

        // Fill the table with numeric values (3 rows x 3 columns).
        int number = 1;
        for (int row = 0; row < 3; row++)
        {
            for (int col = 0; col < 3; col++)
            {
                builder.InsertCell();
                builder.Writeln(number.ToString());
                number++;
            }
            builder.EndRow();
        }

        // End the table.
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Iterate through each cell and apply shading if the numeric value is greater than 5.
        foreach (Row r in table.Rows)
        {
            foreach (Cell c in r.Cells)
            {
                // Get the cell text and trim the end-of-cell marker.
                string cellText = c.GetText().Trim();

                if (int.TryParse(cellText, out int cellValue) && cellValue > 5)
                {
                    // Apply solid yellow shading to the cell.
                    c.CellFormat.Shading.ForegroundPatternColor = Color.Yellow;
                    c.CellFormat.Shading.Texture = TextureIndex.TextureSolid;
                }
            }
        }

        // Save the document.
        string outputPath = "ConditionalShading.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception("The output document was not created.");
        }
    }
}
