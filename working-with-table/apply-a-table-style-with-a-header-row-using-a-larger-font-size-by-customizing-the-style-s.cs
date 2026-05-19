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

        // Start building a table.
        Table table = builder.StartTable();

        // ----- Header row (first row) -----
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.InsertCell();
        builder.Write("Header 3");
        builder.EndRow();

        // ----- Data rows -----
        for (int i = 1; i <= 3; i++)
        {
            builder.InsertCell();
            builder.Write($"Row {i} Col 1");
            builder.InsertCell();
            builder.Write($"Row {i} Col 2");
            builder.InsertCell();
            builder.Write($"Row {i} Col 3");
            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Create a custom table style.
        TableStyle tableStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyHeaderStyle");

        // Increase the font size for the first row (header) via conditional style.
        tableStyle.ConditionalStyles[ConditionalStyleType.FirstRow].Font.Size = 16;

        // Optional: give the header a light gray background.
        tableStyle.ConditionalStyles[ConditionalStyleType.FirstRow].Shading.BackgroundPatternColor = Color.LightGray;

        // Apply the style to the table.
        table.Style = tableStyle;

        // Enable the first‑row conditional formatting.
        table.StyleOptions = TableStyleOptions.FirstRow;

        // Save the document to the local file system.
        string outputPath = "TableStyleHeader.docx";
        doc.Save(outputPath);
    }
}
