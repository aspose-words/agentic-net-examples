using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table.
        Table table = builder.StartTable();

        // Build a simple table with 5 rows and 3 columns.
        for (int row = 0; row < 5; row++)
        {
            for (int col = 0; col < 3; col++)
            {
                builder.InsertCell();
                builder.Write($"Row {row + 1}, Col {col + 1}");
            }
            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Create a custom table style.
        TableStyle tableStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "AlternatingRowStyle");

        // Define the banding to alternate every row.
        tableStyle.RowStripe = 1; // 1 means each consecutive row is a new band.

        // Set shading colors for odd and even rows.
        tableStyle.ConditionalStyles[ConditionalStyleType.OddRowBanding].Shading.BackgroundPatternColor = Color.LightBlue;
        tableStyle.ConditionalStyles[ConditionalStyleType.EvenRowBanding].Shading.BackgroundPatternColor = Color.LightGray;

        // Optional: define a simple border for the table.
        tableStyle.Borders.Color = Color.Black;
        tableStyle.Borders.LineStyle = LineStyle.Single;

        // Apply the custom style to the table.
        table.Style = tableStyle;

        // Enable row banding for the table.
        table.StyleOptions = TableStyleOptions.RowBands;

        // Save the document to the local file system.
        doc.Save("AlternatingRowsTable.docx");
    }
}
