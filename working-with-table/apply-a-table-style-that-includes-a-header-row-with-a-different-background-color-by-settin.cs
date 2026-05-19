using System;
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

        // Start a table.
        Table table = builder.StartTable();

        // ----- Header row -----
        // Insert first header cell.
        builder.InsertCell();
        builder.Write("Product");
        // Insert second header cell.
        builder.InsertCell();
        builder.Write("Quantity");
        // Finish the header row.
        builder.EndRow();

        // ----- Data rows -----
        // First data row.
        builder.InsertCell();
        builder.Write("Apples");
        builder.InsertCell();
        builder.Write("30");
        builder.EndRow();

        // Second data row.
        builder.InsertCell();
        builder.Write("Bananas");
        builder.InsertCell();
        builder.Write("45");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Create a custom table style.
        TableStyle tableStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "HeaderRowStyle");

        // Set a different background color for the header row (first row).
        tableStyle.ConditionalStyles[ConditionalStyleType.FirstRow].Shading.BackgroundPatternColor = Color.LightBlue;

        // Apply the style to the table.
        table.Style = tableStyle;

        // Enable the first row conditional formatting so the header style is applied.
        table.StyleOptions = TableStyleOptions.FirstRow;

        // Save the document to a file.
        string outputPath = "HeaderRowStyleTable.docx";
        doc.Save(outputPath);
    }
}
