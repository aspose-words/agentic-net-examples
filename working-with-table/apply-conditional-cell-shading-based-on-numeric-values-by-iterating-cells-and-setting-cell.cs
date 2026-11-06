using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple table with numeric values.
        builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Write("5");
        builder.InsertCell();
        builder.Write("15");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Write("20");
        builder.InsertCell();
        builder.Write("8");
        builder.EndRow();

        builder.EndTable();

        // Save the initial document.
        string initialPath = "ConditionalShading.docx";
        doc.Save(initialPath);

        // Apply conditional shading: cells with a numeric value greater than 12 get a light green background.
        foreach (Table table in doc.GetChildNodes(NodeType.Table, true))
        {
            foreach (Row row in table.Rows)
            {
                foreach (Cell cell in row.Cells)
                {
                    // Extract the cell text and try to parse it as an integer.
                    string cellText = cell.ToString(SaveFormat.Text).Trim();
                    if (int.TryParse(cellText, out int numericValue))
                    {
                        if (numericValue > 12)
                        {
                            // Apply solid shading with a light green background.
                            cell.CellFormat.Shading.Texture = TextureIndex.TextureSolid;
                            cell.CellFormat.Shading.BackgroundPatternColor = Color.LightGreen;
                        }
                    }
                }
            }
        }

        // Save the document with shading applied.
        string shadedPath = "ConditionalShading_Shaded.docx";
        doc.Save(shadedPath);

        // Verify that the output file exists.
        if (!File.Exists(shadedPath))
        {
            throw new Exception("The shaded document was not saved correctly.");
        }
    }
}
