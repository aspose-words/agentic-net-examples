using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;

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
        builder.InsertCell();
        builder.Write("Product");
        builder.InsertCell();
        builder.Write("Price");
        builder.EndRow();

        // ----- Data rows -----
        builder.InsertCell();
        builder.Write("Apple");
        builder.InsertCell();
        builder.Write("$1");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("Banana");
        builder.InsertCell();
        builder.Write("$2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Create a custom table style.
        TableStyle tableStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "HeaderHighlightStyle");

        // Set the background color for the first (header) row via conditional style.
        tableStyle.ConditionalStyles[ConditionalStyleType.FirstRow].Shading.BackgroundPatternColor = Color.LightBlue;

        // Apply the style to the table.
        table.Style = tableStyle;

        // Enable the first‑row conditional formatting.
        table.StyleOptions = TableStyleOptions.FirstRow;

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableStyleHeaderRow.docx");
        doc.Save(outputPath);

        // Simple validation to ensure the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved correctly.");
    }
}
