using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a 3x3 table.
        Table table = builder.StartTable();
        for (int row = 0; row < 3; row++)
        {
            for (int col = 0; col < 3; col++)
            {
                builder.InsertCell();
                builder.Write($"R{row + 1}C{col + 1}");
            }
            builder.EndRow();
        }
        builder.EndTable();

        // Create a custom table style that will contain a conditional style for the last column.
        TableStyle tableStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyConditionalColumnStyle");

        // Apply bold font and a background color to the last column via its conditional style.
        tableStyle.ConditionalStyles.LastColumn.Font.Bold = true;
        tableStyle.ConditionalStyles.LastColumn.Shading.BackgroundPatternColor = Color.LightYellow;

        // Assign the custom style to the table.
        table.Style = tableStyle;

        // Enable the conditional formatting for the last column.
        table.StyleOptions |= TableStyleOptions.LastColumn;

        // Save the document as HTML, exporting only relative widths for tables, rows and cells.
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            TableWidthOutputMode = HtmlElementSizeOutputMode.RelativeOnly
        };
        doc.Save("ConditionalColumn.html", htmlOptions);
    }
}
