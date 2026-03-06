using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a 3x3 table.
        builder.StartTable();
        for (int row = 0; row < 3; row++)
        {
            for (int col = 0; col < 3; col++)
            {
                builder.InsertCell();
                builder.Write($"R{row + 1}C{col + 1}");
            }
            builder.EndRow();
        }
        builder.EndTable();

        // Create a custom table style.
        TableStyle tableStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyColumnStyle");

        // Apply conditional formatting to the first column: make text bold and set a background color.
        tableStyle.ConditionalStyles.FirstColumn.Font.Bold = true;
        tableStyle.ConditionalStyles.FirstColumn.Shading.BackgroundPatternColor = Color.LightYellow;

        // Retrieve the created table and assign the custom style.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        table.Style = tableStyle;

        // Enable the first‑column conditional style for this table.
        table.StyleOptions |= TableStyleOptions.FirstColumn;

        // Save the document as a Word file.
        doc.Save("TableWithColumnConditional.docx");

        // Export the same document to Markdown, keeping the table as raw HTML.
        MarkdownSaveOptions mdOptions = new MarkdownSaveOptions
        {
            ExportAsHtml = MarkdownExportAsHtml.Tables
        };
        doc.Save("TableWithColumnConditional.md", mdOptions);
    }
}
