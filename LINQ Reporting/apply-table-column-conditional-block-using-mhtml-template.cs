using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Tables;

class TableConditionalColumnMhtml
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table and add a few rows with two columns each.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.EndRow();

        // Data rows.
        for (int i = 1; i <= 3; i++)
        {
            builder.InsertCell();
            builder.Write($"Row {i} Col 1");
            builder.InsertCell();
            builder.Write($"Row {i} Col 2");
            builder.EndRow();
        }

        builder.EndTable();

        // Create a custom table style.
        TableStyle customStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyCustomTableStyle");

        // Apply a conditional style to the first column – set a light gray background.
        customStyle.ConditionalStyles.FirstColumn.Shading.BackgroundPatternColor = Color.LightGray;

        // Optionally change the font of the first column to bold.
        customStyle.ConditionalStyles.FirstColumn.Font.Bold = true;

        // Assign the custom style to the table.
        table.Style = customStyle;

        // Enable the first‑column conditional formatting for this table.
        table.StyleOptions |= TableStyleOptions.FirstColumn;

        // Save the document as MHTML, preserving relative widths only.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
        {
            TableWidthOutputMode = HtmlElementSizeOutputMode.RelativeOnly
        };

        doc.Save("TableConditionalColumn.mht", saveOptions);
    }
}
