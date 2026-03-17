using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;

class AlternatingRowColors
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table.
        Table table = builder.StartTable();

        // ----- Header row -----
        // Apply a distinct background to the header via the style later (FirstRow conditional style).
        builder.InsertCell();
        builder.Write("Product");
        builder.InsertCell();
        builder.Write("Quantity");
        builder.EndRow();

        // ----- Data rows -----
        // Add several rows; the style will shade odd/even rows automatically.
        for (int i = 1; i <= 6; i++)
        {
            builder.InsertCell();
            builder.Write($"Item {i}");
            builder.InsertCell();
            builder.Write((i * 10).ToString());
            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // ----- Create a custom table style -----
        TableStyle tableStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "AlternatingStyle");

        // Enable row banding (alternating colors) and first‑row formatting.
        tableStyle.RowStripe = 1; // Alternate after each row.
        table.StyleOptions = TableStyleOptions.RowBands | TableStyleOptions.FirstRow;

        // Header (first row) background.
        tableStyle.ConditionalStyles[ConditionalStyleType.FirstRow].Shading.BackgroundPatternColor = Color.LightGray;

        // Odd rows background.
        tableStyle.ConditionalStyles[ConditionalStyleType.OddRowBanding].Shading.BackgroundPatternColor = Color.LightBlue;

        // Even rows background.
        tableStyle.ConditionalStyles[ConditionalStyleType.EvenRowBanding].Shading.BackgroundPatternColor = Color.LightCyan;

        // Apply the style to the table.
        table.Style = tableStyle;

        // Save the document.
        doc.Save("AlternatingRows.docx");
    }
}
