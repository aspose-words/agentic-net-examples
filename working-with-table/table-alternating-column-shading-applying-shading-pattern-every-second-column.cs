using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;

class AlternatingColumnShading
{
    static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple table with 4 columns and 5 rows.
        Table table = builder.StartTable();
        for (int row = 0; row < 5; row++)
        {
            for (int col = 0; col < 4; col++)
            {
                builder.InsertCell();
                builder.Writeln($"Row {row + 1}, Col {col + 1}");
            }
            builder.EndRow();
        }
        builder.EndTable();

        // Create a custom table style.
        TableStyle tableStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "AlternatingColumnStyle");

        // Apply a background color to every even column.
        tableStyle.ColumnStripe = 1; // Apply the color to each column in the band.
        tableStyle.ConditionalStyles[ConditionalStyleType.EvenColumnBanding]
                  .Shading.BackgroundPatternColor = Color.LightSalmon;

        // Assign the style to the table.
        table.Style = tableStyle;

        // Enable column banding for the table.
        table.StyleOptions |= TableStyleOptions.ColumnBands;

        // Save the document.
        string outputPath = "AlternatingColumnShading.docx";
        doc.Save(outputPath);
    }
}
