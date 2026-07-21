using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 2‑column table with a few rows.
        Table table = builder.StartTable();

        for (int i = 0; i < 6; i++)
        {
            builder.InsertCell();
            builder.Write($"Row {i + 1}, Column 1");
            builder.InsertCell();
            builder.Write($"Row {i + 1}, Column 2");
            builder.EndRow();
        }

        builder.EndTable();

        // Create a custom table style that will alternate row shading.
        TableStyle alternatingStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "AlternatingRowStyle");
        alternatingStyle.RowStripe = 1; // Apply banding to each row.

        // Define shading for odd and even rows.
        alternatingStyle.ConditionalStyles[ConditionalStyleType.OddRowBanding].Shading.BackgroundPatternColor = Color.LightBlue;
        alternatingStyle.ConditionalStyles[ConditionalStyleType.EvenRowBanding].Shading.BackgroundPatternColor = Color.LightGray;

        // Apply the style to the table and enable row banding.
        table.Style = alternatingStyle;
        table.StyleOptions = TableStyleOptions.RowBands;

        // Save the document.
        doc.Save("AlternatingRowsTable.docx");
    }
}
