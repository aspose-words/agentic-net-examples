using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;

class TableAlternatingRows
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple table with a few rows.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Write("Item");
        builder.InsertCell();
        builder.Write("Quantity");
        builder.EndRow();

        // Data rows.
        string[] items = { "Apples", "Bananas", "Carrots" };
        int[] qty = { 20, 40, 50 };
        for (int i = 0; i < items.Length; i++)
        {
            builder.InsertCell();
            builder.Write(items[i]);
            builder.InsertCell();
            builder.Write(qty[i].ToString());
            builder.EndRow();
        }

        builder.EndTable();

        // Create a custom table style.
        TableStyle tableStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyAlternatingRowStyle");

        // Define the colors for odd and even row banding.
        tableStyle.ConditionalStyles[ConditionalStyleType.OddRowBanding].Shading.BackgroundPatternColor = Color.LightBlue;
        tableStyle.ConditionalStyles[ConditionalStyleType.EvenRowBanding].Shading.BackgroundPatternColor = Color.LightCyan;

        // Apply the style to the table.
        table.Style = tableStyle;

        // Enable row banding (alternating row shading) for this table.
        table.StyleOptions = table.StyleOptions | TableStyleOptions.RowBands;

        // Save the document.
        doc.Save("TableAlternatingRows.docx");
    }
}
