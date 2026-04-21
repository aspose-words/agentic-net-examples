using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 2x2 table.
        Table table = builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Write("Name");
        builder.InsertCell();
        builder.Write("Value");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Write("Alice");
        builder.InsertCell();
        builder.Write("42");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Create a custom table style.
        TableStyle customStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyTableStyle1");
        customStyle.Shading.BackgroundPatternColor = Color.LightBlue;
        customStyle.Borders.Color = Color.DarkBlue;
        customStyle.Borders.LineStyle = LineStyle.Single;

        // Apply the style to the table via the Table.Style property.
        table.Style = customStyle;

        // Save the document.
        string outputPath = "TableStyleExample.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output file was not created.");
    }
}
