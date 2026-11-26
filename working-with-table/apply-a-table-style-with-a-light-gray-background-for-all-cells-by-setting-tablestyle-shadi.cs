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

        // Build a simple 2x2 table.
        builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Write("Cell 3");
        builder.InsertCell();
        builder.Write("Cell 4");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Create a new table style.
        Style grayStyle = doc.Styles.Add(StyleType.Table, "GrayTableStyle");

        // Cast to TableStyle to access shading properties.
        TableStyle grayTableStyle = (TableStyle)grayStyle;

        // Set a solid light gray background for all cells.
        grayTableStyle.Shading.BackgroundPatternColor = Color.LightGray;
        grayTableStyle.Shading.Texture = TextureIndex.TextureNone;

        // Apply the style to the table.
        table.Style = grayTableStyle;

        // Save the document.
        string outputPath = "TableStyleShading.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception("The output document was not created.");
        }
    }
}
