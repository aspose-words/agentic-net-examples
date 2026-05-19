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

        // Create a custom table style named "CustomStyle".
        TableStyle customStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "CustomStyle");

        // Define basic style properties.
        customStyle.AllowBreakAcrossPages = true;
        customStyle.CellSpacing = 4;               // Space between cells.
        customStyle.BottomPadding = 5;
        customStyle.TopPadding = 5;
        customStyle.LeftPadding = 5;
        customStyle.RightPadding = 5;

        // Set shading (background color) for the style.
        customStyle.Shading.BackgroundPatternColor = Color.LightYellow;

        // Set default borders for the style.
        customStyle.Borders.Color = Color.DarkBlue;
        customStyle.Borders.LineStyle = LineStyle.Single;
        customStyle.Borders.LineWidth = 1.0; // 1 point thickness.

        // Apply the custom style to the table.
        table.Style = customStyle;

        // Save the document to a file in the current directory.
        string outputPath = "CustomTableStyle.docx";
        doc.Save(outputPath);
    }
}
