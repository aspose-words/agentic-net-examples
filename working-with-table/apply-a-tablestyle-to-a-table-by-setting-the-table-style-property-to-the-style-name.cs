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

        // First row.
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndTable(); // Ends the table and returns the Table node.

        // Create a custom table style.
        TableStyle customStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyCustomTableStyle");
        customStyle.Shading.BackgroundPatternColor = Color.LightGray;
        customStyle.Borders.Color = Color.DarkBlue;
        customStyle.Borders.LineStyle = LineStyle.Single;
        customStyle.Borders.LineWidth = 1.0;
        customStyle.Alignment = TableAlignment.Center;

        // Apply the style to the table by assigning the Style property.
        // Retrieve the style by name from the document's style collection.
        table.Style = doc.Styles["MyCustomTableStyle"];

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableStyleExample.docx");
        doc.Save(outputPath);

        // Simple validation to ensure the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The document was not saved correctly.");
    }
}
