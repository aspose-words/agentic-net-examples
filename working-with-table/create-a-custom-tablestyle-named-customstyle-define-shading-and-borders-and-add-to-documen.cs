using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a custom table style and add it to the document's style collection.
        TableStyle customStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "CustomStyle");

        // Define shading.
        customStyle.Shading.BackgroundPatternColor = Color.LightYellow;

        // Define borders.
        customStyle.Borders.Color = Color.DarkBlue;
        customStyle.Borders.LineWidth = 2.0;
        customStyle.Borders.LineStyle = LineStyle.Single;

        // Build a simple 2x2 table.
        builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Writeln("Cell 1");
        builder.InsertCell();
        builder.Writeln("Cell 2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Writeln("Cell 3");
        builder.InsertCell();
        builder.Writeln("Cell 4");
        builder.EndRow();

        builder.EndTable();

        // Apply the custom style to the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        table.Style = customStyle;

        // Save the document.
        string outputPath = "CustomStyleTable.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception($"Failed to create the output file: {outputPath}");
        }
    }
}
