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

        // Add a custom table style named "CustomStyle".
        // The style itself does not expose Shading or Borders directly,
        // so we will apply the formatting to the table after it is created.
        Style customStyle = doc.Styles.Add(StyleType.Table, "CustomStyle");

        // Build a simple 2x1 table.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Apply the custom style to the table.
        table.Style = customStyle;

        // Define shading for each cell in the table (light blue background).
        foreach (Row row in table.Rows)
        {
            foreach (Cell cell in row.Cells)
            {
                cell.CellFormat.Shading.BackgroundPatternColor = Color.LightBlue;
                cell.CellFormat.Shading.Texture = TextureIndex.TextureNone;
            }
        }

        // Define borders for the table (single black line).
        table.SetBorder(BorderType.Left, LineStyle.Single, 1.0, Color.Black, true);
        table.SetBorder(BorderType.Right, LineStyle.Single, 1.0, Color.Black, true);
        table.SetBorder(BorderType.Top, LineStyle.Single, 1.0, Color.Black, true);
        table.SetBorder(BorderType.Bottom, LineStyle.Single, 1.0, Color.Black, true);
        table.SetBorder(BorderType.Horizontal, LineStyle.Single, 1.0, Color.Black, true);
        table.SetBorder(BorderType.Vertical, LineStyle.Single, 1.0, Color.Black, true);

        // Save the document.
        string outputPath = "CustomTableStyle.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved correctly.");
    }
}
