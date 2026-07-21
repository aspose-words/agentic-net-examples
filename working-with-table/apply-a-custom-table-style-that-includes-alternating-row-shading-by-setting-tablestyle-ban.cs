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
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 4‑row table.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Write("Item");
        builder.InsertCell();
        builder.Write("Quantity");
        builder.EndRow();

        // Data rows.
        string[] items = { "Apples", "Bananas", "Carrots", "Dates" };
        int[] qty = { 20, 40, 50, 60 };
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
        TableStyle customStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyAlternatingRowStyle");
        // Enable row banding (alternating shading).
        customStyle.RowStripe = 1; // Apply shading every row.
        // Define shading for odd rows.
        customStyle.ConditionalStyles[ConditionalStyleType.OddRowBanding].Shading.BackgroundPatternColor = Color.LightBlue;
        // Define shading for even rows.
        customStyle.ConditionalStyles[ConditionalStyleType.EvenRowBanding].Shading.BackgroundPatternColor = Color.LightGray;

        // Apply the style to the table.
        table.Style = customStyle;
        // Enable the row banding option.
        table.StyleOptions = TableStyleOptions.RowBands;

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "CustomTableStyle.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved successfully.");
    }
}
