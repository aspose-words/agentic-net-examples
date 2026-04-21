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
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("Cell 3");
        builder.InsertCell();
        builder.Write("Cell 4");
        builder.EndTable();

        // Create a custom table style and apply it to the table.
        TableStyle customStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyCustomStyle");
        customStyle.CellSpacing = 2.0;
        customStyle.Shading.BackgroundPatternColor = Color.LightGray;
        customStyle.Borders.Color = Color.Blue;
        customStyle.Borders.LineStyle = LineStyle.Single;
        table.Style = customStyle;

        // Iterate through all styles in the document and modify table style definitions.
        foreach (Style style in doc.Styles)
        {
            if (style.Type == StyleType.Table)
            {
                TableStyle tableStyle = (TableStyle)style;
                // Example modifications: increase cell spacing, change shading and border color.
                tableStyle.CellSpacing = 5.0;
                tableStyle.Shading.BackgroundPatternColor = Color.Yellow;
                tableStyle.Borders.Color = Color.Red;
                tableStyle.Borders.LineStyle = LineStyle.DotDash;
            }
        }

        // Convert any remaining style formatting to direct formatting (optional but ensures changes are visible).
        doc.ExpandTableStylesToDirectFormatting();

        // Save the resulting document.
        doc.Save("UpdatedTableStyle.docx");
    }
}
