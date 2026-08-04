using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 2x2 table.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndTable();

        // Create a custom table style and set its properties.
        TableStyle customStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyCustomStyle");
        customStyle.CellSpacing = 5;
        customStyle.BottomPadding = 10;
        customStyle.LeftPadding = 5;
        customStyle.RightPadding = 5;
        customStyle.TopPadding = 10;
        customStyle.Shading.BackgroundPatternColor = Color.AntiqueWhite;
        customStyle.Borders.Color = Color.Blue;
        customStyle.Borders.LineStyle = LineStyle.DotDash;
        customStyle.VerticalAlignment = CellVerticalAlignment.Center;

        // Apply the custom style to the table.
        table.Style = customStyle;

        // Expand the style into direct formatting on the table, rows, and cells.
        doc.ExpandTableStylesToDirectFormatting();

        // Save the resulting document.
        string fileName = "TableStyleToDirectFormatting.docx";
        doc.Save(fileName);

        // Verify that the file was saved successfully.
        if (!File.Exists(fileName))
            throw new Exception("Document was not saved correctly.");

        // Inform the user where the file was saved.
        Console.WriteLine($"Document saved to: {Path.GetFullPath(fileName)}");
    }
}
