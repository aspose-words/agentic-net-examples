using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class TableConditionalStyleExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 5‑row table.
        Table table = builder.StartTable();
        for (int r = 1; r <= 5; r++)
        {
            for (int c = 1; c <= 3; c++)
            {
                builder.InsertCell();
                builder.Write($"R{r}C{c}");
            }
            builder.EndRow();
        }
        builder.EndTable();

        // -----------------------------------------------------------------
        // Create a custom table style that will contain conditional formatting.
        // -----------------------------------------------------------------
        TableStyle customStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyCustomTableStyle");

        // Example: apply a background color to the last row of the table.
        customStyle.ConditionalStyles[ConditionalStyleType.LastRow].Shading.BackgroundPatternColor = System.Drawing.Color.LightGreen;

        // Example: apply bold font to the first column of the table.
        customStyle.ConditionalStyles[ConditionalStyleType.FirstColumn].Font.Bold = true;

        // Example: apply alternating row banding (odd rows light blue, even rows light cyan).
        customStyle.RowStripe = 1; // band every single row
        customStyle.ConditionalStyles[ConditionalStyleType.OddRowBanding].Shading.BackgroundPatternColor = System.Drawing.Color.LightBlue;
        customStyle.ConditionalStyles[ConditionalStyleType.EvenRowBanding].Shading.BackgroundPatternColor = System.Drawing.Color.LightCyan;

        // Assign the custom style to the table.
        table.Style = customStyle;

        // Enable the style options that we have defined.
        // FirstRow, FirstColumn and RowBands are enabled by default for a new style.
        // We need to explicitly enable LastRow and ColumnBands if we want them applied.
        table.StyleOptions |= TableStyleOptions.LastRow | TableStyleOptions.ColumnBands;

        // -----------------------------------------------------------------
        // Save the document.
        // -----------------------------------------------------------------
        doc.Save("TableWithConditionalStyles.docx", SaveFormat.Docx);
    }
}
