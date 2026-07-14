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

        // Build a simple 2x2 table.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();
        builder.EndTable();

        // Create a custom table style and apply initial formatting.
        TableStyle customStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyCustomTableStyle");
        customStyle.CellSpacing = 2;
        customStyle.Shading.BackgroundPatternColor = Color.LightGray;
        customStyle.Borders.Color = Color.DarkBlue;
        customStyle.Borders.LineStyle = LineStyle.Single;

        // Assign the custom style to the table.
        table.Style = customStyle;

        // Iterate through all styles in the document.
        foreach (Style style in doc.Styles)
        {
            // Process only table styles.
            if (style.Type == StyleType.Table)
            {
                TableStyle tableStyle = (TableStyle)style;

                // Update style properties programmatically.
                tableStyle.Shading.BackgroundPatternColor = Color.LightYellow;
                tableStyle.CellSpacing = 5;
                tableStyle.Borders.Color = Color.Green;
            }
        }

        // Convert style formatting to direct formatting on tables.
        doc.ExpandTableStylesToDirectFormatting();

        // Save the resulting document.
        doc.Save("UpdatedTableStyles.docx");
    }
}
