using System;
using System.IO;
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

        // Build a simple table with a header row and a few data rows.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Write("Item");
        builder.InsertCell();
        builder.Write("Quantity");
        builder.EndRow();

        // Data rows.
        string[] items = { "Apples", "Bananas", "Carrots", "Dates" };
        int[] quantities = { 10, 20, 30, 40 };
        for (int i = 0; i < items.Length; i++)
        {
            builder.InsertCell();
            builder.Write(items[i]);
            builder.InsertCell();
            builder.Write(quantities[i].ToString());
            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Create a custom table style that uses row banding (alternating shading).
        TableStyle customStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyAlternatingStyle");
        // One row per band creates the classic alternating pattern.
        customStyle.RowStripe = 1;
        // Define shading for odd rows.
        customStyle.ConditionalStyles[ConditionalStyleType.OddRowBanding].Shading.BackgroundPatternColor = Color.LightGray;
        // Define shading for even rows.
        customStyle.ConditionalStyles[ConditionalStyleType.EvenRowBanding].Shading.BackgroundPatternColor = Color.White;

        // Apply the custom style to the table.
        table.Style = customStyle;
        // Enable row banding via the style options flag.
        table.StyleOptions = TableStyleOptions.RowBands;

        // Save the document to the current directory.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "AlternatingRows.docx");
        doc.Save(outputPath);
    }
}
