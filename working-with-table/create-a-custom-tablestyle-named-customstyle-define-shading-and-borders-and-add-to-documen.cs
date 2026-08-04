using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using System.Drawing;

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
        builder.Write("Cell 1,1");
        builder.InsertCell();
        builder.Write("Cell 1,2");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("Cell 2,1");
        builder.InsertCell();
        builder.Write("Cell 2,2");
        builder.EndTable();

        // Create a custom table style named "CustomStyle".
        TableStyle customStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "CustomStyle");

        // Define shading (background color) for the style.
        customStyle.Shading.BackgroundPatternColor = Color.LightBlue;

        // Define borders for the style.
        customStyle.Borders.Color = Color.DarkBlue;
        customStyle.Borders.LineStyle = LineStyle.Single;
        customStyle.Borders.LineWidth = 1.5; // Optional: set border thickness.

        // Optionally set some padding.
        customStyle.LeftPadding = 5;
        customStyle.RightPadding = 5;
        customStyle.TopPadding = 5;
        customStyle.BottomPadding = 5;

        // Apply the custom style to the table.
        table.Style = customStyle;

        // Verify that the style was applied (optional).
        if (table.StyleName != "CustomStyle")
            throw new InvalidOperationException("Custom style was not applied to the table.");

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(outputDir);

        // Save the document.
        string outputPath = Path.Combine(outputDir, "CustomTableStyle.docx");
        doc.Save(outputPath);
    }
}
