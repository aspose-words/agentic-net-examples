using System;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Tables;   // Needed for the Table class

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
        builder.EndTable(); // Ends the table and moves the cursor out.

        // Apply a single black border to all sides of the table.
        table.SetBorders(LineStyle.Single, 1.0, Color.Black);

        // Prepare HTML save options – export only the table, no headers/footers.
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            ExportHeadersFootersMode = ExportHeadersFootersMode.None
        };

        // Export the table to an HTML fragment (preserves borders).
        string htmlFragment = table.ToString(htmlOptions);

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(outputDir);

        // Save the HTML fragment to a file.
        string htmlPath = Path.Combine(outputDir, "TableFragment.html");
        File.WriteAllText(htmlPath, htmlFragment);

        // Indicate completion.
        Console.WriteLine($"HTML fragment saved to: {htmlPath}");
    }
}
