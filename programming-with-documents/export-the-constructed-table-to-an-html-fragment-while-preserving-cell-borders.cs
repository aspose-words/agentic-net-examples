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

        // Apply a solid black border to the whole table.
        table.SetBorders(LineStyle.Single, 1.0, Color.Black);

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

        // Export the table node to an HTML fragment.
        string htmlFragment = table.ToString(SaveFormat.Html);

        // Save the HTML fragment to a file.
        string fragmentPath = Path.Combine(Directory.GetCurrentDirectory(), "TableFragment.html");
        File.WriteAllText(fragmentPath, htmlFragment);

        // (Optional) Save the whole document as HTML for visual verification.
        string docHtmlPath = Path.Combine(Directory.GetCurrentDirectory(), "Document.html");
        doc.Save(docHtmlPath, SaveFormat.Html);
    }
}
