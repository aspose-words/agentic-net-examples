using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table.
        Table table = builder.StartTable();

        // Build a simple 2‑row, 4‑column table.
        for (int row = 0; row < 2; row++)
        {
            for (int col = 0; col < 4; col++)
            {
                builder.InsertCell();
                builder.Writeln($"R{row + 1}C{col + 1}");
            }
            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Create a custom table style that will apply alternating column shading.
        TableStyle customStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyColumnBandStyle");
        // Apply banding to every column.
        customStyle.ColumnStripe = 1;
        // Define shading for even columns.
        customStyle.ConditionalStyles[ConditionalStyleType.EvenColumnBanding].Shading.BackgroundPatternColor = Color.LightGray;
        // Define shading for odd columns.
        customStyle.ConditionalStyles[ConditionalStyleType.OddColumnBanding].Shading.BackgroundPatternColor = Color.White;

        // Assign the custom style to the table.
        table.Style = customStyle;
        // Enable column banding.
        table.StyleOptions = TableStyleOptions.ColumnBands;

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableColumnBanding.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The document was not saved correctly.");
    }
}
