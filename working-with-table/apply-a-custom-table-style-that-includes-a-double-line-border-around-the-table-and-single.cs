using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class TableStyleExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a 3x3 table.
        builder.StartTable();
        for (int row = 0; row < 3; row++)
        {
            for (int col = 0; col < 3; col++)
            {
                builder.InsertCell();
                builder.Writeln($"R{row}C{col}");
            }
            builder.EndRow();
        }
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChildNodes(NodeType.Table, true)[0];

        // Apply custom borders:
        // - Double line border around the outer edges of the table.
        // - Single line borders for internal cell dividers.
        for (int i = 0; i < table.Rows.Count; i++)
        {
            Row row = table.Rows[i];
            for (int j = 0; j < row.Cells.Count; j++)
            {
                Cell cell = row.Cells[j];
                BorderCollection borders = cell.CellFormat.Borders;

                bool isOuter = i == 0 || i == table.Rows.Count - 1 || j == 0 || j == row.Cells.Count - 1;

                LineStyle style = isOuter ? LineStyle.Double : LineStyle.Single;
                borders[BorderType.Left].LineStyle = style;
                borders[BorderType.Right].LineStyle = style;
                borders[BorderType.Top].LineStyle = style;
                borders[BorderType.Bottom].LineStyle = style;
            }
        }

        // Save the document.
        string outputPath = "TableStyleExample.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not saved correctly.");
    }
}
