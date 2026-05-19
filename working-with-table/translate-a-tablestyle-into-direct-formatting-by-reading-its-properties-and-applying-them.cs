using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 2x2 table.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("Cell 3");
        builder.InsertCell();
        builder.Write("Cell 4");
        builder.EndTable();

        // Create a custom table style and set several style properties.
        TableStyle customStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyCustomStyle");
        customStyle.CellSpacing = 5;                                   // Space between cells.
        customStyle.Shading.BackgroundPatternColor = Color.AntiqueWhite; // Cell background color.
        customStyle.Borders.Color = Color.Blue;                         // Border color.
        customStyle.Borders.LineStyle = LineStyle.DotDash;              // Border line style.
        customStyle.RowStripe = 2;                                      // Number of rows in a band.
        customStyle.ColumnStripe = 2;                                   // Number of columns in a band.
        customStyle.Alignment = TableAlignment.Center;                  // Table alignment.

        // Apply the style to the table.
        table.Style = customStyle;

        // Convert the style formatting into direct formatting on the table elements.
        doc.ExpandTableStylesToDirectFormatting();

        // Verify that some of the style properties have been transferred to direct formatting.
        Console.WriteLine("Table CellSpacing (direct): " + table.CellSpacing);
        Console.WriteLine("First cell background color (direct): " +
            table.FirstRow.FirstCell.CellFormat.Shading.BackgroundPatternColor.Name);
        Console.WriteLine("First cell border color (direct): " +
            table.FirstRow.FirstCell.CellFormat.Borders[BorderType.Left].Color.Name);
        Console.WriteLine("Table alignment (direct): " + table.Alignment);

        // Save the resulting document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableStyleToDirectFormatting.docx");
        doc.Save(outputPath);
        Console.WriteLine("Document saved to: " + outputPath);
    }
}
