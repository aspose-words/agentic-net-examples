using System;
using System.IO;
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

        // Add a custom table style.
        Style customStyle = doc.Styles.Add(StyleType.Table, "MyCustomTableStyle");
        // Set an initial shading for the style (will be modified later).
        TableStyle tableStyle = (TableStyle)customStyle;
        tableStyle.Shading.BackgroundPatternColor = Color.LightGray;

        // Build a simple 2x2 table and apply the custom style.
        builder.StartTable();
        builder.RowFormat.Height = 20;
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();
        builder.InsertCell();
        builder.Write("Cell 3");
        builder.InsertCell();
        builder.Write("Cell 4");
        builder.EndRow();
        builder.EndTable();

        // Retrieve the created table and assign the custom style.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        table.Style = customStyle; // assign the Style object, not its name

        // Iterate through all styles in the document.
        foreach (Style style in doc.Styles)
        {
            // Process only table styles.
            if (style.Type == StyleType.Table)
            {
                TableStyle ts = (TableStyle)style;

                // Example modification: change shading background to LightBlue.
                ts.Shading.BackgroundPatternColor = Color.LightBlue;

                // Example modification: set all borders to a solid red line.
                foreach (Border border in ts.Borders)
                {
                    border.LineStyle = LineStyle.Single;
                    border.Color = Color.Red;
                    border.LineWidth = 1.0; // points
                }

                // RowBandSize and ColumnBandSize are not available in this version,
                // so they are omitted.
            }
        }

        // Save the document.
        string outputPath = "UpdatedTableStyle.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception("The output document was not saved correctly.");
        }
    }
}
