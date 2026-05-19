using System;
using Aspose.Words;
using Aspose.Words.Tables;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table.
        Table table = builder.StartTable();

        // Populate the table with sample data (6 rows, 2 columns).
        for (int i = 0; i < 6; i++)
        {
            builder.InsertCell();
            builder.Write($"Row {i + 1}, Cell 1");
            builder.InsertCell();
            builder.Write($"Row {i + 1}, Cell 2");
            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Create a custom table style that uses row banding.
        TableStyle tableStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "AlternatingRowStyle");
        // Alternate shading every row.
        tableStyle.RowStripe = 1;
        // Define shading for odd rows.
        tableStyle.ConditionalStyles[ConditionalStyleType.OddRowBanding].Shading.BackgroundPatternColor = Color.LightBlue;
        // Define shading for even rows.
        tableStyle.ConditionalStyles[ConditionalStyleType.EvenRowBanding].Shading.BackgroundPatternColor = Color.LightGray;

        // Apply the style to the table.
        table.Style = tableStyle;
        // Enable row banding for the table.
        table.StyleOptions = TableStyleOptions.RowBands;

        // Save the document to disk.
        doc.Save("AlternatingRows.docx");
    }
}
