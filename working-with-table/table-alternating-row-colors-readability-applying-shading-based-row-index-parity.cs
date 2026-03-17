using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;

class AlternatingRowTable
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple table with a few rows and columns.
        Table table = builder.StartTable();
        for (int i = 0; i < 6; i++) // 6 rows
        {
            for (int j = 0; j < 3; j++) // 3 columns
            {
                builder.InsertCell();
                builder.Write($"Row {i + 1}, Col {j + 1}");
            }
            builder.EndRow();
        }
        builder.EndTable();

        // Create a custom table style that will apply alternating row shading.
        TableStyle alternatingStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "AlternatingRows");

        // Enable row banding (alternating rows) and set the stripe size to 1 (every row).
        alternatingStyle.RowStripe = 1; // one row per band

        // Define the background colors for odd and even rows.
        alternatingStyle.ConditionalStyles[ConditionalStyleType.OddRowBanding]
            .Shading.BackgroundPatternColor = Color.LightBlue;
        alternatingStyle.ConditionalStyles[ConditionalStyleType.EvenRowBanding]
            .Shading.BackgroundPatternColor = Color.LightCyan;

        // Optional: set a simple border for the table.
        alternatingStyle.Borders.Color = Color.Black;
        alternatingStyle.Borders.LineStyle = LineStyle.Single;

        // Apply the style to the table and enable row banding.
        table.Style = alternatingStyle;
        table.StyleOptions = TableStyleOptions.RowBands;

        // Save the document.
        doc.Save("AlternatingRows.docx");
    }
}
