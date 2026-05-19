using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 2‑column table with a header row and several data rows.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Writeln("Item");
        builder.InsertCell();
        builder.Writeln("Quantity");
        builder.EndRow();

        // Data rows.
        for (int i = 1; i <= 5; i++)
        {
            builder.InsertCell();
            builder.Writeln($"Item {i}");
            builder.InsertCell();
            builder.Writeln((i * 10).ToString());
            builder.EndRow();
        }

        builder.EndTable();

        // Create a custom table style.
        TableStyle customStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "AlternatingRows");
        // Define the banding interval (1 = every row).
        customStyle.RowStripe = 1;
        // Set shading for odd rows.
        customStyle.ConditionalStyles[ConditionalStyleType.OddRowBanding].Shading.BackgroundPatternColor = Color.LightGray;
        // Set shading for even rows.
        customStyle.ConditionalStyles[ConditionalStyleType.EvenRowBanding].Shading.BackgroundPatternColor = Color.White;

        // Apply the style to the table and enable row banding.
        table.Style = customStyle;
        table.StyleOptions = TableStyleOptions.RowBands;

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "AlternatingRowsTable.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved correctly.");
    }
}
