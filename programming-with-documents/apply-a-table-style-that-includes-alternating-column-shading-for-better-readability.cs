using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "TableWithAlternatingColumnShading.docx");

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple table (3 rows x 4 columns).
        Table table = builder.StartTable();
        for (int row = 0; row < 3; row++)
        {
            for (int col = 0; col < 4; col++)
            {
                builder.InsertCell();
                builder.Write($"R{row + 1}C{col + 1}");
            }
            builder.EndRow();
        }
        builder.EndTable();

        // Create a custom table style that defines alternating column shading.
        TableStyle tableStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "AlternatingColumnStyle");
        // Apply shading to even columns.
        tableStyle.ConditionalStyles[ConditionalStyleType.EvenColumnBanding].Shading.BackgroundPatternColor = Color.LightGray;
        // Apply shading to odd columns (optional, can be left as default white).
        tableStyle.ConditionalStyles[ConditionalStyleType.OddColumnBanding].Shading.BackgroundPatternColor = Color.White;
        // Set the banding to alternate every column.
        tableStyle.ColumnStripe = 1;

        // Apply the style to the table.
        table.Style = tableStyle;
        // Enable column banding for this table.
        table.StyleOptions = TableStyleOptions.ColumnBands;

        // Save the document.
        doc.Save(outputPath);
    }
}
