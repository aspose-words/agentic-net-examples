using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;
using System.Drawing;

class ConditionalRowMarkdownExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 3‑row, 2‑column table.
        Table table = builder.StartTable();

        // Row 1
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.EndRow();

        // Row 2
        builder.InsertCell();
        builder.Write("Data A1");
        builder.InsertCell();
        builder.Write("Data A2");
        builder.EndRow();

        // Row 3
        builder.InsertCell();
        builder.Write("Data B1");
        builder.InsertCell();
        builder.Write("Data B2");
        builder.EndTable();

        // Create a custom table style.
        TableStyle customStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyCustomStyle");

        // Apply a conditional style to the first row (e.g., background shading).
        customStyle.ConditionalStyles[ConditionalStyleType.FirstRow].Shading.BackgroundPatternColor = Color.LightBlue;
        customStyle.ConditionalStyles[ConditionalStyleType.FirstRow].Font.Bold = true;
        customStyle.ConditionalStyles[ConditionalStyleType.FirstRow].ParagraphFormat.Alignment = ParagraphAlignment.Center;

        // Enable the FirstRow conditional formatting for the table.
        table.Style = customStyle;
        table.StyleOptions = TableStyleOptions.FirstRow;

        // Export the document to Markdown.
        MarkdownSaveOptions mdOptions = new MarkdownSaveOptions
        {
            // Export tables as native Markdown (no raw HTML).
            ExportAsHtml = MarkdownExportAsHtml.None
        };

        // Save the result as a .md file.
        doc.Save("ConditionalRowTable.md", mdOptions);
    }
}
