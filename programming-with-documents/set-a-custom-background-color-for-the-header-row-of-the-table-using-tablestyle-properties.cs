using System;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple table with a header row and two data rows.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Write("Product");
        builder.InsertCell();
        builder.Write("Quantity");
        builder.EndRow();

        // First data row.
        builder.InsertCell();
        builder.Write("Apples");
        builder.InsertCell();
        builder.Write("10");
        builder.EndRow();

        // Second data row.
        builder.InsertCell();
        builder.Write("Bananas");
        builder.InsertCell();
        builder.Write("20");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Create a custom table style.
        TableStyle headerStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "HeaderRowStyle");

        // Set the background color for the first row (header) using conditional style.
        headerStyle.ConditionalStyles[ConditionalStyleType.FirstRow].Shading.BackgroundPatternColor = Color.LightBlue;

        // Apply the style to the table.
        table.Style = headerStyle;

        // Ensure the style is applied to the first row.
        table.StyleOptions = TableStyleOptions.FirstRow;

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "HeaderRowStyle.docx");
        doc.Save(outputPath);
    }
}
