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

        // Add a custom table style named "CustomStyle".
        TableStyle customStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "CustomStyle");

        // Define shading for the style (solid light blue background).
        customStyle.Shading.Texture = TextureIndex.TextureSolid;
        customStyle.Shading.BackgroundPatternColor = Color.LightBlue;

        // Define borders for the style (single dark blue borders).
        foreach (Border border in customStyle.Borders)
        {
            border.LineStyle = LineStyle.Single;
            border.Color = Color.DarkBlue;
            border.LineWidth = 1.0; // points
        }

        // Build a simple table and apply the custom style.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.EndRow();
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();
        builder.EndTable();

        // Retrieve the created table and assign the custom style.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        table.Style = customStyle;

        // Save the document.
        string outputPath = "CustomTableStyle.docx";
        doc.Save(outputPath);
    }
}
