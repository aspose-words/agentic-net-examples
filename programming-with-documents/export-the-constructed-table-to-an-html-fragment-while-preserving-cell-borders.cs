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

        // Start a table.
        Table table = builder.StartTable();

        // Apply a single line border to all sides of the table.
        table.SetBorders(LineStyle.Single, 1.0, Color.Black);

        // First row (header).
        builder.InsertCell();
        builder.Writeln("Header 1");
        builder.InsertCell();
        builder.Writeln("Header 2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Writeln("Cell A1");
        builder.InsertCell();
        builder.Writeln("Cell A2");
        builder.EndRow();

        // Third row.
        builder.InsertCell();
        builder.Writeln("Cell B1");
        builder.InsertCell();
        builder.Writeln("Cell B2");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Define the output HTML file path.
        string outputPath = "TableFragment.html";

        // Configure HTML save options to produce a fragment (no full HTML document) and preserve borders.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);
        saveOptions.PrettyFormat = true;
        saveOptions.ExportImagesAsBase64 = true;
        // The default behavior already omits headers/footers and produces a fragment,
        // so no additional settings are required for this simple example.

        // Save the document as an HTML fragment.
        doc.Save(outputPath, saveOptions);

        // Verify that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"HTML fragment saved successfully to '{outputPath}'.");
        }
        else
        {
            Console.WriteLine("Failed to save the HTML fragment.");
        }
    }
}
