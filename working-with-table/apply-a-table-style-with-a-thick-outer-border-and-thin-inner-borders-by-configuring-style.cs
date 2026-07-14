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
        Table table = builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Write("Cell 1,1");
        builder.InsertCell();
        builder.Write("Cell 1,2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Write("Cell 2,1");
        builder.InsertCell();
        builder.Write("Cell 2,2");
        builder.EndTable();

        // Create a custom table style.
        TableStyle tableStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyTableStyle");
        // Configure the style's default cell borders (inner borders) to be thin.
        tableStyle.Borders.LineStyle = LineStyle.Single;
        tableStyle.Borders.LineWidth = 0.5; // thin inner borders
        tableStyle.Borders.Color = Color.Black;

        // Apply the style to the table.
        table.Style = tableStyle;

        // Set thick outer borders directly on the table.
        table.SetBorder(BorderType.Left,   LineStyle.Single, 2.0, Color.Black, false);
        table.SetBorder(BorderType.Right,  LineStyle.Single, 2.0, Color.Black, false);
        table.SetBorder(BorderType.Top,    LineStyle.Single, 2.0, Color.Black, false);
        table.SetBorder(BorderType.Bottom, LineStyle.Single, 2.0, Color.Black, false);

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableWithStyle.docx");
        doc.Save(outputPath);

        // Simple verification that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The document was not saved correctly.");
    }
}
