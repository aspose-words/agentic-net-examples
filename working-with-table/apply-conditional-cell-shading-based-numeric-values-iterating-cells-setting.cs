using Aspose.Words;
using Aspose.Words.Tables;
using System.Drawing;

class Program
{
    static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a table with numeric values.
        builder.StartTable();
        for (int i = 0; i < 5; i++)
        {
            for (int j = 0; j < 3; j++)
            {
                builder.InsertCell();
                int value = (i + 1) * (j + 1);
                builder.Write(value.ToString());
            }
            builder.EndRow();
        }
        builder.EndTable();

        // Get the first table in the document.
        Table table = doc.FirstSection.Body.Tables[0];

        // Iterate over each cell, parse its numeric content, and apply shading.
        foreach (Row row in table.Rows)
        {
            foreach (Cell cell in row.Cells)
            {
                // Cell text includes a cell marker (\\a); trim it before parsing.
                string cellText = cell.GetText().Trim('\a', '\r', '\n');
                if (int.TryParse(cellText, out int number))
                {
                    // Example rule: numbers > 6 get a light green background,
                    // otherwise a light salmon background.
                    if (number > 6)
                        cell.CellFormat.Shading.BackgroundPatternColor = Color.LightGreen;
                    else
                        cell.CellFormat.Shading.BackgroundPatternColor = Color.LightSalmon;
                }
            }
        }

        // Save the resulting document.
        doc.Save("ConditionalShadingTable.docx");
    }
}
