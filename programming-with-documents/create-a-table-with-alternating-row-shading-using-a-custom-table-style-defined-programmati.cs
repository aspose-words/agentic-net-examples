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

        // Build a simple 5‑row, 2‑column table.
        Table table = builder.StartTable();
        for (int row = 1; row <= 5; row++)
        {
            // First cell.
            builder.InsertCell();
            builder.Write($"Row {row}, Cell 1");

            // Second cell.
            builder.InsertCell();
            builder.Write($"Row {row}, Cell 2");

            // End the current row.
            builder.EndRow();
        }
        builder.EndTable();

        // Create a custom table style that alternates row shading.
        TableStyle alternatingStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "AlternatingRowStyle");
        // Apply banding to each row (alternates every row).
        alternatingStyle.RowStripe = 1;
        // Define shading colors for odd and even rows.
        alternatingStyle.ConditionalStyles[ConditionalStyleType.OddRowBanding].Shading.BackgroundPatternColor = Color.LightBlue;
        alternatingStyle.ConditionalStyles[ConditionalStyleType.EvenRowBanding].Shading.BackgroundPatternColor = Color.LightCyan;

        // Assign the style to the table and enable row banding.
        table.Style = alternatingStyle;
        table.StyleOptions = TableStyleOptions.RowBands;

        // Save the document to the current directory.
        doc.Save("AlternatingRows.docx");
    }
}
