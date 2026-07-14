using System;
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

        // -----------------------------------------------------------------
        // Create a custom table style.
        // -----------------------------------------------------------------
        TableStyle customStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyHeaderStyle");

        // Increase the font size for the first row (header) of the table.
        // ConditionalStyles[ConditionalStyleType.FirstRow] targets the header row.
        customStyle.ConditionalStyles[ConditionalStyleType.FirstRow].Font.Size = 20;

        // Optionally set a background color for the header row to make it stand out.
        customStyle.ConditionalStyles[ConditionalStyleType.FirstRow].Shading.BackgroundPatternColor = System.Drawing.Color.LightGray;

        // -----------------------------------------------------------------
        // Build a simple table with a header row.
        // -----------------------------------------------------------------
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Write("Product");
        builder.InsertCell();
        builder.Write("Quantity");
        builder.EndRow();

        // Data rows.
        builder.InsertCell();
        builder.Write("Apples");
        builder.InsertCell();
        builder.Write("50");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("Bananas");
        builder.InsertCell();
        builder.Write("30");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Apply the custom style to the table.
        table.Style = customStyle;

        // Enable the FirstRow style option so that the conditional formatting is applied.
        table.StyleOptions = TableStyleOptions.FirstRow;

        // -----------------------------------------------------------------
        // Save the document to the local file system.
        // -----------------------------------------------------------------
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableWithHeaderStyle.docx");
        doc.Save(outputPath);
    }
}
