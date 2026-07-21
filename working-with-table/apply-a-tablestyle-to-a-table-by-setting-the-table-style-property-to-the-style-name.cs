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
        Table table = builder.StartTable();

        // First row (header).
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.EndRow();

        // Second row (data).
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Create a custom table style.
        TableStyle customStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyCustomTableStyle");
        customStyle.Shading.BackgroundPatternColor = Color.LightGray;
        customStyle.Borders.Color = Color.Blue;
        customStyle.Borders.LineStyle = LineStyle.Single;
        customStyle.Borders.LineWidth = 1.5;
        customStyle.CellSpacing = 2;
        customStyle.BottomPadding = 5;
        customStyle.TopPadding = 5;
        customStyle.LeftPadding = 5;
        customStyle.RightPadding = 5;

        // Apply the style to the table using the Table.Style property.
        table.Style = customStyle;

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableStyleExample.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The document was not saved correctly.");

        // Indicate successful completion.
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
