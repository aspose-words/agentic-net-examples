using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table.
        builder.StartTable();

        // First row – header cells.
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.EndRow();

        // Second row – data cells.
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Retrieve the created table to apply borders.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        table.SetBorder(BorderType.Left, LineStyle.Single, 1.0, Color.Black, true);
        table.SetBorder(BorderType.Right, LineStyle.Single, 1.0, Color.Black, true);
        table.SetBorder(BorderType.Top, LineStyle.Single, 1.0, Color.Black, true);
        table.SetBorder(BorderType.Bottom, LineStyle.Single, 1.0, Color.Black, true);
        table.SetBorder(BorderType.Horizontal, LineStyle.Single, 1.0, Color.Black, true);
        table.SetBorder(BorderType.Vertical, LineStyle.Single, 1.0, Color.Black, true);

        // Configure HTML save options to produce a fragment (no full HTML document wrapper).
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            ExportHeadersFootersMode = ExportHeadersFootersMode.None,
            PrettyFormat = true,
            ExportImagesAsBase64 = true
        };

        // Save the table as an HTML fragment.
        const string outputPath = "TableFragment.html";
        doc.Save(outputPath, saveOptions);
    }
}
