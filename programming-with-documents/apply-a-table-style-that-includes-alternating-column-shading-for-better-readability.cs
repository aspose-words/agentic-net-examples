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

        // Build a simple table with 4 columns and 5 rows.
        Table table = builder.StartTable();
        for (int row = 1; row <= 5; row++)
        {
            for (int col = 1; col <= 4; col++)
            {
                builder.InsertCell();
                builder.Write($"R{row}C{col}");
            }
            builder.EndRow();
        }
        builder.EndTable();

        // Create a custom table style that will shade columns alternately.
        TableStyle tableStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "AlternatingColumnStyle");
        // Each band consists of a single column.
        tableStyle.ColumnStripe = 1;
        // Define shading for odd columns.
        tableStyle.ConditionalStyles[ConditionalStyleType.OddColumnBanding].Shading.BackgroundPatternColor = Color.LightGray;
        // Define shading for even columns.
        tableStyle.ConditionalStyles[ConditionalStyleType.EvenColumnBanding].Shading.BackgroundPatternColor = Color.White;

        // Apply the style to the table.
        table.Style = tableStyle;
        // Enable column banding for the table.
        table.StyleOptions = TableStyleOptions.ColumnBands;

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(outputDir);

        // Save the document.
        string outputPath = Path.Combine(outputDir, "AlternatingColumnShading.docx");
        doc.Save(outputPath);
    }
}
