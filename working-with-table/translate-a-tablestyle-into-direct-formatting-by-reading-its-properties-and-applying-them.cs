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
        builder.EndTable();

        // Create a custom table style and set some formatting properties.
        TableStyle customStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyCustomStyle");
        customStyle.CellSpacing = 4; // Space between cells.
        customStyle.Shading.BackgroundPatternColor = Color.LightBlue; // Cell background.
        customStyle.Borders.Color = Color.DarkBlue; // Border color.
        customStyle.Borders.LineStyle = LineStyle.Single; // Border line style.

        // Apply the style to the table.
        table.Style = customStyle;

        // Convert the style's formatting into direct formatting on the table elements.
        doc.ExpandTableStylesToDirectFormatting();

        // Simple validation that the style was expanded.
        if (Math.Abs(table.CellSpacing - 4) > 0.001)
            throw new Exception("CellSpacing was not applied as direct formatting.");

        // Save the resulting document.
        doc.Save("TableStyleToDirectFormatting.docx");
    }
}
