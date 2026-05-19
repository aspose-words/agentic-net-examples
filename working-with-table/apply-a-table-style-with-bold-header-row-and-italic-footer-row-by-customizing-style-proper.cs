using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;

public class TableStyleExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 3‑row table: header, data row, footer.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Write("Header");
        builder.InsertCell();
        builder.Write("Column 2");
        builder.EndRow();

        // Data row.
        builder.InsertCell();
        builder.Write("Data 1");
        builder.InsertCell();
        builder.Write("Data 2");
        builder.EndRow();

        // Footer row.
        builder.InsertCell();
        builder.Write("Footer");
        builder.InsertCell();
        builder.Write("Summary");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Create a custom table style.
        TableStyle tableStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyCustomTableStyle");

        // Make the first row (header) bold.
        tableStyle.ConditionalStyles[ConditionalStyleType.FirstRow].Font.Bold = true;

        // Make the last row (footer) italic.
        tableStyle.ConditionalStyles[ConditionalStyleType.LastRow].Font.Italic = true;

        // Apply the style to the table.
        table.Style = tableStyle;

        // Enable the conditional formatting for first and last rows.
        table.StyleOptions = TableStyleOptions.FirstRow | TableStyleOptions.LastRow;

        // Save the document.
        string outputPath = "TableStyleExample.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception($"Failed to create the output file: {outputPath}");
    }
}
