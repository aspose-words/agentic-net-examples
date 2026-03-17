using Aspose.Words;
using Aspose.Words.Tables;
using System.Drawing;

class Program
{
    static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 5‑row, 4‑column table.
        Table table = builder.StartTable();
        for (int row = 0; row < 5; row++)
        {
            for (int col = 0; col < 4; col++)
            {
                builder.InsertCell();
                builder.Write($"R{row + 1}C{col + 1}");
            }
            builder.EndRow();
        }
        builder.EndTable();

        // Create a custom table style.
        TableStyle tableStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "AltColumnStyle");

        // Enable column banding: apply shading to every even column.
        tableStyle.ColumnStripe = 1; // Band every column.
        tableStyle.ConditionalStyles[ConditionalStyleType.EvenColumnBanding].Shading.BackgroundPatternColor = Color.LightSalmon;

        // Assign the style to the table.
        table.Style = tableStyle;

        // Turn on column banding for this table.
        table.StyleOptions |= TableStyleOptions.ColumnBands;

        // Save the document.
        doc.Save("AlternatingColumnShading.docx");
    }
}
