using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;
using System.Drawing;

class TableColumnConditionalExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 3x3 table.
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

        // Create a custom table style that will hold the conditional formatting.
        TableStyle customStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyCustomTableStyle");

        // Apply a background color to the first column via its conditional style.
        // This will affect only the cells that belong to the first column of any table using this style.
        customStyle.ConditionalStyles.FirstColumn.Shading.BackgroundPatternColor = Color.LightGreen;

        // Optionally, make the text in the first column bold.
        customStyle.ConditionalStyles.FirstColumn.Font.Bold = true;

        // Enable the first‑column conditional formatting for the table.
        // The FirstColumn flag is on by default for most built‑in styles, but we set it explicitly.
        table.StyleOptions = TableStyleOptions.FirstColumn | TableStyleOptions.ColumnBands;

        // Assign the custom style to the table.
        table.Style = customStyle;

        // Save the document.
        doc.Save("TableColumnConditional.docx");
    }
}
