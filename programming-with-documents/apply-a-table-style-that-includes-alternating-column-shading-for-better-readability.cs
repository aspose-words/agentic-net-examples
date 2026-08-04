using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class ApplyAlternatingColumnShading
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table and add a header row.
        Table table = builder.StartTable();

        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.InsertCell();
        builder.Write("Header 3");
        builder.EndRow();

        // Add a few data rows.
        for (int i = 1; i <= 5; i++)
        {
            builder.InsertCell();
            builder.Write($"Row {i} Col 1");
            builder.InsertCell();
            builder.Write($"Row {i} Col 2");
            builder.InsertCell();
            builder.Write($"Row {i} Col 3");
            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Create a custom table style that will shade columns alternately.
        TableStyle tableStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "AlternatingColumnStyle");

        // Define the number of columns that make up a band (1 = every column).
        tableStyle.ColumnStripe = 1;

        // Set shading for odd columns.
        tableStyle.ConditionalStyles[ConditionalStyleType.OddColumnBanding]
                  .Shading.BackgroundPatternColor = Color.LightBlue;

        // Set shading for even columns.
        tableStyle.ConditionalStyles[ConditionalStyleType.EvenColumnBanding]
                  .Shading.BackgroundPatternColor = Color.LightGray;

        // Apply the style to the table.
        table.Style = tableStyle;

        // Enable column banding for the table.
        table.StyleOptions = TableStyleOptions.ColumnBands;

        // Auto‑fit the table to its contents.
        table.AutoFit(AutoFitBehavior.AutoFitToContents);

        // Save the document.
        string outputPath = "AlternatingColumnShading.docx";
        doc.Save(outputPath);
    }
}
