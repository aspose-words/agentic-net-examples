using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple table.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Header");
        builder.InsertCell();
        builder.Write("Value");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("Item 1");
        builder.InsertCell();
        builder.Write("10");
        builder.EndRow();

        builder.EndTable();

        // Create a custom table style.
        TableStyle tableStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyCustomTableStyle");
        tableStyle.Shading.BackgroundPatternColor = Color.LightYellow;
        tableStyle.Borders.Color = Color.DarkBlue;
        tableStyle.Borders.LineStyle = LineStyle.Single;
        tableStyle.Borders.LineWidth = 1.5;

        // Apply the style to the table via the Style property.
        table.Style = tableStyle;

        // Save the document.
        string outputPath = "TableStyleExample.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Failed to create the output document.");
    }
}
