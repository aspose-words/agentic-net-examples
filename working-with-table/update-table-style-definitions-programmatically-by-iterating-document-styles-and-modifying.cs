using System;
using System.Drawing;
using System.IO;
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
        builder.StartTable();
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

        // Apply an existing table style to the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        table.Style = doc.Styles["Table Grid"];

        // Iterate through all styles in the document.
        foreach (Style style in doc.Styles)
        {
            // Process only table styles.
            if (style.Type == StyleType.Table)
            {
                // Cast to TableStyle to access table‑specific formatting.
                TableStyle tableStyle = (TableStyle)style;

                // Change the background shading of the whole table style.
                tableStyle.Shading.BackgroundPatternColor = Color.LightBlue;

                // Set a uniform border color for the style.
                tableStyle.Borders.Color = Color.DarkBlue;

                // Note: RowBandSize and ColumnBandSize properties are not available
                // in the current Aspose.Words version, so they are omitted.
            }
        }

        // Save the modified document.
        string outputPath = "UpdatedTableStyle.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception($"Failed to create the output file: {outputPath}");
        }
    }
}
