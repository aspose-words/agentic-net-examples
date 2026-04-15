using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 2x2 table.
        builder.StartTable();

        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("Cell 3");
        builder.InsertCell();
        builder.Write("Cell 4");
        builder.EndTable();

        // Retrieve the created table.
        Table table = doc.FirstSection.Body.Tables[0];

        // Apply borders to all sides of the table and its cells.
        table.SetBorder(BorderType.Left, LineStyle.Single, 1.0, Color.Black, true);
        table.SetBorder(BorderType.Right, LineStyle.Single, 1.0, Color.Black, true);
        table.SetBorder(BorderType.Top, LineStyle.Single, 1.0, Color.Black, true);
        table.SetBorder(BorderType.Bottom, LineStyle.Single, 1.0, Color.Black, true);
        table.SetBorder(BorderType.Horizontal, LineStyle.Single, 1.0, Color.Black, true);
        table.SetBorder(BorderType.Vertical, LineStyle.Single, 1.0, Color.Black, true);

        // Export the table to an HTML fragment, preserving the borders.
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
        string htmlFragment = table.ToString(htmlOptions);

        // Write the HTML fragment to a file.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableFragment.html");
        File.WriteAllText(outputPath, htmlFragment);

        // Optionally, display the path of the generated file.
        Console.WriteLine($"HTML fragment saved to: {outputPath}");
    }
}
