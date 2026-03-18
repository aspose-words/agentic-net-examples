using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using System.Drawing;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 2‑column table with a header row.
        Table table = builder.StartTable();

        // Header cells.
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.EndRow();

        // Data row.
        builder.InsertCell();
        builder.Write("Data 1");
        builder.InsertCell();
        builder.Write("Data 2");
        builder.EndRow();

        builder.EndTable();

        // Create a custom table style.
        TableStyle tableStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyHeaderStyle");

        // Increase the font size for the first row (the header) via its conditional style.
        tableStyle.ConditionalStyles[ConditionalStyleType.FirstRow].Font.Size = 16; // larger font size

        // Optional: give the header a background colour for visual distinction.
        tableStyle.ConditionalStyles[ConditionalStyleType.FirstRow].Shading.BackgroundPatternColor = Color.LightGray;

        // Apply the custom style to the table.
        table.Style = tableStyle;

        // Enable the first‑row conditional formatting so the style takes effect.
        table.StyleOptions = TableStyleOptions.FirstRow;

        // Save the document.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "TableStyleHeader.docx");
        doc.Save(outputPath);
    }
}
