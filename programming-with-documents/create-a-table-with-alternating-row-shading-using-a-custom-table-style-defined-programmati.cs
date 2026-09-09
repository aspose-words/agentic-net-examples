using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 5‑row, 2‑column table.
        Table table = builder.StartTable();
        for (int i = 1; i <= 5; i++)
        {
            // First column.
            builder.InsertCell();
            builder.Write($"Row {i}, Col 1");

            // Second column.
            builder.InsertCell();
            builder.Write($"Row {i}, Col 2");

            builder.EndRow();
        }
        builder.EndTable();

        // Create a custom table style.
        TableStyle customStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "AlternatingRowStyle");

        // Define how many rows participate in the banding (1 = every row).
        customStyle.RowStripe = 1;

        // Set shading for odd rows.
        customStyle.ConditionalStyles[ConditionalStyleType.OddRowBanding].Shading.BackgroundPatternColor = Color.LightBlue;

        // Set shading for even rows.
        customStyle.ConditionalStyles[ConditionalStyleType.EvenRowBanding].Shading.BackgroundPatternColor = Color.LightCyan;

        // Apply the style to the table.
        table.Style = customStyle;

        // Enable row banding for the table.
        table.StyleOptions |= TableStyleOptions.RowBands;

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "AlternatingRowsTable.docx");
        doc.Save(outputPath);
    }
}
