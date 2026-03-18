using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

class ConditionalEvenRowsExample
{
    static void Main()
    {
        // Folder to save the document.
        string artifactsDir = @"C:\Temp\"; // Adjust path as needed.

        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple table with 6 rows and 2 columns.
        Table table = builder.StartTable();
        for (int i = 0; i < 6; i++)
        {
            for (int j = 0; j < 2; j++)
            {
                builder.InsertCell();
                builder.Write($"Row {i + 1}, Col {j + 1}");
            }
            builder.EndRow();
        }
        builder.EndTable();

        // Create a custom table style.
        TableStyle evenRowStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "EvenRowStyle");

        // Define the banding interval (1 = every row).
        evenRowStyle.RowStripe = 1;

        // Apply shading to even‑indexed rows via the conditional style.
        evenRowStyle.ConditionalStyles[ConditionalStyleType.EvenRowBanding]
                     .Shading.BackgroundPatternColor = Color.LightGray;

        // Assign the style to the table.
        table.Style = evenRowStyle;

        // Enable row banding so the conditional style takes effect.
        table.StyleOptions |= TableStyleOptions.RowBands;

        // Save the document.
        doc.Save(artifactsDir + "ConditionalEvenRows.docx");
    }
}
