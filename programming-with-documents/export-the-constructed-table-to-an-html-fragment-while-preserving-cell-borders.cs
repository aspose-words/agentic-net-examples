using System;
using System.IO;
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

        // Apply borders to the whole table (all sides and inner borders).
        table.SetBorders(LineStyle.Single, 1.0, Color.Black);

        // Save the document as an HTML fragment (body only, no wrapper).
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            ExportHeadersFootersMode = ExportHeadersFootersMode.None
            // No need to set ExportEmbeddedCss – it does not exist on HtmlSaveOptions.
        };

        // Determine output path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableFragment.html");

        // Save the document.
        doc.Save(outputPath, saveOptions);

        // Inform the user.
        Console.WriteLine($"HTML fragment saved to: {outputPath}");
    }
}
