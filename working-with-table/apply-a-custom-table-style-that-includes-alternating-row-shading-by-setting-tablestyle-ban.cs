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

        // Build a simple table with a header row and a few data rows.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Write("Item");
        builder.InsertCell();
        builder.Write("Quantity");
        builder.EndRow();

        // Data rows.
        string[] items = { "Apples", "Bananas", "Carrots" };
        int[] quantities = { 20, 40, 50 };

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

        // Create a custom table style.
        TableStyle customStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyBandingStyle");

        // Define the banding interval (alternating rows).
        customStyle.RowStripe = 1; // Alternate every row.

        // Set shading for odd rows.
        customStyle.ConditionalStyles[ConditionalStyleType.OddRowBanding].Shading.BackgroundPatternColor = Color.LightGray;

        // Set shading for even rows.
        customStyle.ConditionalStyles[ConditionalStyleType.EvenRowBanding].Shading.BackgroundPatternColor = Color.White;

        // Apply the custom style to the table.
        table.Style = customStyle;

        // Enable row banding for the table.
        table.StyleOptions = TableStyleOptions.RowBands;

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableWithAlternatingRowShading.docx");
        doc.Save(outputPath);
    }
}
