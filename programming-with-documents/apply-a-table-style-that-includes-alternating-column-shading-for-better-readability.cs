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

        // Start a table with 4 columns and 5 rows (including header).
        Table table = builder.StartTable();

        // Header row.
        for (int c = 0; c < 4; c++)
        {
            builder.InsertCell();
            builder.Write($"Header {c + 1}");
        }
        builder.EndRow();

        // Data rows.
        for (int r = 0; r < 4; r++)
        {
            for (int c = 0; c < 4; c++)
            {
                builder.InsertCell();
                builder.Write($"R{r + 1}C{c + 1}");
            }
            builder.EndRow();
        }

        builder.EndTable();

        // Create a custom table style that applies alternating column shading.
        TableStyle tableStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "AlternatingColumnStyle");

        // Define the number of columns to include in the banding (1 = every column).
        tableStyle.ColumnStripe = 1;

        // Set shading colors for odd and even column banding.
        tableStyle.ConditionalStyles[ConditionalStyleType.OddColumnBanding].Shading.BackgroundPatternColor = Color.LightBlue;
        tableStyle.ConditionalStyles[ConditionalStyleType.EvenColumnBanding].Shading.BackgroundPatternColor = Color.LightGray;

        // Apply the style to the table.
        table.Style = tableStyle;

        // Enable column banding via StyleOptions.
        table.StyleOptions = table.StyleOptions | TableStyleOptions.ColumnBands;

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(outputDir);

        // Save the document.
        string outputPath = Path.Combine(outputDir, "TableWithAlternatingColumnShading.docx");
        doc.Save(outputPath);
    }
}
