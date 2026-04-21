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

        // Build a simple 3x3 table.
        Table table = builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Write("R1C1");
        builder.InsertCell();
        builder.Write("R1C2");
        builder.InsertCell();
        builder.Write("R1C3");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Write("R2C1");
        builder.InsertCell();
        builder.Write("R2C2");
        builder.InsertCell();
        builder.Write("R2C3");
        builder.EndRow();

        // Third row.
        builder.InsertCell();
        builder.Write("R3C1");
        builder.InsertCell();
        builder.Write("R3C2");
        builder.InsertCell();
        builder.Write("R3C3");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Create a custom table style.
        TableStyle tableStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyCustomTableStyle");

        // Configure the style's default cell borders (inner borders) to be thin.
        tableStyle.Borders.LineStyle = LineStyle.Single;
        tableStyle.Borders.LineWidth = 0.5; // Thin inner borders.
        tableStyle.Borders.Color = Color.Black;

        // Apply the custom style to the table.
        table.Style = tableStyle;

        // Apply thick outer borders directly to the table.
        table.SetBorder(BorderType.Left, LineStyle.Single, 2.0, Color.Black, true);
        table.SetBorder(BorderType.Right, LineStyle.Single, 2.0, Color.Black, true);
        table.SetBorder(BorderType.Top, LineStyle.Single, 2.0, Color.Black, true);
        table.SetBorder(BorderType.Bottom, LineStyle.Single, 2.0, Color.Black, true);

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableStyleBorders.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not saved correctly.");
    }
}
