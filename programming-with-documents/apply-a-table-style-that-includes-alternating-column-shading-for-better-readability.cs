using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;
using Aspose.Words.Themes;
using Aspose.Words.Saving;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table and add a few rows and columns.
        Table table = builder.StartTable();

        // Create header row.
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.InsertCell();
        builder.Write("Header 3");
        builder.EndRow();

        // Add sample data rows.
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

        // End the table construction.
        builder.EndTable();

        // Create a custom table style.
        TableStyle tableStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "AlternatingColumnStyle");

        // Define colors for odd and even column banding.
        tableStyle.ColumnStripe = 1; // Apply banding to each column.
        tableStyle.ConditionalStyles[ConditionalStyleType.OddColumnBanding].Shading.BackgroundPatternColor = Color.LightBlue;
        tableStyle.ConditionalStyles[ConditionalStyleType.EvenColumnBanding].Shading.BackgroundPatternColor = Color.LightSalmon;

        // Apply the style to the table.
        table.Style = tableStyle;

        // Enable column banding for the table.
        table.StyleOptions = TableStyleOptions.ColumnBands;

        // Save the document to the output file.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "AlternatingColumnShading.docx");
        doc.Save(outputPath);
    }
}
