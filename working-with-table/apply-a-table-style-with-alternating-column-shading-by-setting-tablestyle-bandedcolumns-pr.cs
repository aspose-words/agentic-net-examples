using System;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 2‑row, 4‑column table.
        Table table = builder.StartTable();

        // First row (header).
        for (int i = 0; i < 4; i++)
        {
            builder.InsertCell();
            builder.Write($"Header {i + 1}");
        }
        builder.EndRow();

        // Second row (data).
        for (int i = 0; i < 4; i++)
        {
            builder.InsertCell();
            builder.Write($"Data {i + 1}");
        }
        builder.EndRow();

        builder.EndTable();

        // Create a custom table style that defines alternating column shading.
        TableStyle columnBandStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyColumnBandStyle");

        // Define shading for even columns.
        columnBandStyle.ConditionalStyles[ConditionalStyleType.EvenColumnBanding]
            .Shading.BackgroundPatternColor = Color.LightGray;

        // Define shading for odd columns.
        columnBandStyle.ConditionalStyles[ConditionalStyleType.OddColumnBanding]
            .Shading.BackgroundPatternColor = Color.White;

        // Apply the style to the table.
        table.Style = columnBandStyle;

        // Enable column banding for the style.
        table.StyleOptions = TableStyleOptions.ColumnBands;

        // Save the document.
        string outputPath = "TableColumnBanding.docx";
        doc.Save(outputPath);

        // Simple validation to ensure the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
    }
}
