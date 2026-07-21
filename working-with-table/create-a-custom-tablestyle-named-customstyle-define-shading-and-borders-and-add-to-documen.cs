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

        // Define shading for the style.
        customStyle.Shading.BackgroundPatternColor = Color.LightGray;

        // Define borders for the style.
        customStyle.Borders.Color = Color.DarkBlue;
        customStyle.Borders.LineStyle = LineStyle.Single;
        // Optionally set line width for each border side.
        customStyle.Borders.Left.LineWidth = 1.5;
        customStyle.Borders.Right.LineWidth = 1.5;
        customStyle.Borders.Top.LineWidth = 1.5;
        customStyle.Borders.Bottom.LineWidth = 1.5;

        // Apply the custom style to the table.
        table.Style = customStyle;

        // Save the document to a file.
        const string outputPath = "CustomTableStyle.docx";
        doc.Save(outputPath);
    }
}
