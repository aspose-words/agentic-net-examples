using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Tables;

class ExportTableToHtml
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a 2x2 table.
        Table table = builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Write("Cell 3");
        builder.InsertCell();
        builder.Write("Cell 4");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Ensure the table has visible borders.
        table.ClearBorders(); // remove any existing borders
        table.SetBorder(BorderType.Left,   LineStyle.Single, 1.0, Color.Black, true);
        table.SetBorder(BorderType.Right,  LineStyle.Single, 1.0, Color.Black, true);
        table.SetBorder(BorderType.Top,    LineStyle.Single, 1.0, Color.Black, true);
        table.SetBorder(BorderType.Bottom, LineStyle.Single, 1.0, Color.Black, true);
        // Inside borders (optional, but keep cell borders visible).
        table.SetBorder(BorderType.Horizontal, LineStyle.Single, 0.5, Color.Gray, true);
        table.SetBorder(BorderType.Vertical,   LineStyle.Single, 0.5, Color.Gray, true);

        // Configure HTML save options to produce a fragment without extra page/section markup.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            ExportHeadersFootersMode = ExportHeadersFootersMode.None,
            ExportPageSetup = false
        };

        // Export the table node to an HTML string.
        string htmlFragment = table.ToString(saveOptions);

        // Output the HTML fragment.
        Console.WriteLine(htmlFragment);
    }
}
