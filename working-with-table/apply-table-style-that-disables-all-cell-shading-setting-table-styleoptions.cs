using System;
using Aspose.Words;
using Aspose.Words.Tables;
using System.Drawing;

class ApplyNoShading
{
    static void Main()
    {
        // Create a new document and a table.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();
        builder.EndTable();

        // (Optional) Apply shading to each cell to demonstrate removal.
        foreach (Row row in table.Rows)
        {
            foreach (Cell cell in row.Cells)
            {
                cell.CellFormat.Shading.BackgroundPatternColor = Color.Yellow;
            }
        }

        // Remove any shading that may be applied directly to the table.
        table.ClearShading();

        // Ensure that no table style shading is applied.
        table.StyleOptions = TableStyleOptions.None;

        // Save the modified document.
        doc.Save("Table.NoShading.docx");
    }
}
