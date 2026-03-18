using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a 4x4 table.
        Table table = builder.StartTable();
        for (int row = 0; row < 4; row++)
        {
            for (int col = 0; col < 4; col++)
            {
                builder.InsertCell();
                builder.Write($"R{row + 1}C{col + 1}");
            }
            builder.EndRow();
        }
        builder.EndTable();

        // Create a custom table style that will shade alternating columns.
        TableStyle altColumnStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "AltColumnShading");
        // Apply banding to each column (alternates every column).
        altColumnStyle.ColumnStripe = 1;
        // Define shading for even columns.
        altColumnStyle.ConditionalStyles[ConditionalStyleType.EvenColumnBanding].Shading.BackgroundPatternColor = Color.LightGray;
        // Define shading for odd columns (optional, can be left as default).
        altColumnStyle.ConditionalStyles[ConditionalStyleType.OddColumnBanding].Shading.BackgroundPatternColor = Color.White;

        // Assign the style to the table.
        table.Style = altColumnStyle;
        // Enable column banding for this table.
        table.StyleOptions = table.StyleOptions | TableStyleOptions.ColumnBands;

        // Save the document.
        doc.Save("AlternatingColumnShading.docx");
    }
}
