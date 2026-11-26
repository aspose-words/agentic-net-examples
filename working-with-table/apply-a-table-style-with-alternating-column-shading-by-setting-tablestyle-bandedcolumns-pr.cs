using System;
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

        // Obtain a built‑in table style.
        TableStyle style = (TableStyle)doc.Styles["Table Grid"];

        // Apply the style to the table.
        table.Style = style;

        // Enable alternating column shading using the ColumnBands option.
        table.StyleOptions = TableStyleOptions.ColumnBands;

        // Save the document.
        string outputPath = "TableStyleBandedColumns.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not saved correctly.");
    }
}
