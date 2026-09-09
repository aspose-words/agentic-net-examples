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

        // Start a table and add a few rows/columns.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.InsertCell();
        builder.Write("Header 3");
        builder.EndRow();

        // Data rows.
        for (int i = 0; i < 4; i++)
        {
            builder.InsertCell();
            builder.Write($"Row {i + 1} Col 1");
            builder.InsertCell();
            builder.Write($"Row {i + 1} Col 2");
            builder.InsertCell();
            builder.Write($"Row {i + 1} Col 3");
            builder.EndRow();
        }

        builder.EndTable();

        // Create a custom table style.
        TableStyle tableStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "AlternatingColumnStyle");

        // Define shading for odd and even column banding.
        tableStyle.ConditionalStyles[ConditionalStyleType.OddColumnBanding].Shading.BackgroundPatternColor = Color.LightBlue;
        tableStyle.ConditionalStyles[ConditionalStyleType.EvenColumnBanding].Shading.BackgroundPatternColor = Color.LightSalmon;

        // Enable column banding for the table.
        table.Style = tableStyle;
        table.StyleOptions = table.StyleOptions | TableStyleOptions.ColumnBands;

        // Ensure the output directory exists.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "AlternatingColumnTable.docx");
        string outputDir = Path.GetDirectoryName(outputPath);
        if (!Directory.Exists(outputDir))
            Directory.CreateDirectory(outputDir);

        // Save the document.
        doc.Save(outputPath);
    }
}
