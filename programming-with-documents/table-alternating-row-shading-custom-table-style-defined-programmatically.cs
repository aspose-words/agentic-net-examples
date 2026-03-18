using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new blank document and a DocumentBuilder to construct its contents.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple table with 5 rows and 3 columns.
        Table table = builder.StartTable();
        for (int row = 0; row < 5; row++)
        {
            for (int col = 0; col < 3; col++)
            {
                builder.InsertCell();
                builder.Write($"R{row + 1}C{col + 1}");
            }
            builder.EndRow();
        }
        builder.EndTable();

        // Create a custom table style that will provide alternating row shading.
        TableStyle alternatingStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "AlternatingRows");
        // RowStripe = 1 means the banding alternates every single row.
        alternatingStyle.RowStripe = 1;

        // Define shading for odd rows.
        alternatingStyle.ConditionalStyles[ConditionalStyleType.OddRowBanding]
            .Shading.BackgroundPatternColor = Color.LightGray;

        // Define shading for even rows.
        alternatingStyle.ConditionalStyles[ConditionalStyleType.EvenRowBanding]
            .Shading.BackgroundPatternColor = Color.White;

        // Apply the custom style to the table.
        table.Style = alternatingStyle;
        // Enable row banding for this table.
        table.StyleOptions |= TableStyleOptions.RowBands;

        // Save the document to disk.
        doc.Save("AlternatingRowsTable.docx");
    }
}
