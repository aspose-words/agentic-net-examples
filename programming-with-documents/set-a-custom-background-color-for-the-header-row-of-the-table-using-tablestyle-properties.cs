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

        // Build a simple table with a header row and a few data rows.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Writeln("Product");
        builder.InsertCell();
        builder.Writeln("Quantity");
        builder.EndRow();

        // Data rows.
        for (int i = 1; i <= 3; i++)
        {
            builder.InsertCell();
            builder.Writeln($"Item {i}");
            builder.InsertCell();
            builder.Writeln((i * 10).ToString());
            builder.EndRow();
        }

        builder.EndTable();

        // Create a custom table style.
        TableStyle customStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyHeaderStyle");

        // Set the background color for the first row (header) via conditional style.
        customStyle.ConditionalStyles[ConditionalStyleType.FirstRow].Shading.BackgroundPatternColor = Color.LightBlue;

        // Apply the style to the table.
        table.Style = customStyle;

        // Enable the FirstRow conditional formatting.
        table.StyleOptions = TableStyleOptions.FirstRow;

        // Save the document to the output folder.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "HeaderRowStyle.docx");
        doc.Save(outputPath);
    }
}
