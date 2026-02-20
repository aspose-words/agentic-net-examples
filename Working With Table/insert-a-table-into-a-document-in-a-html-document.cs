using System;
using Aspose.Words;
using Aspose.Words.Tables;

class InsertTableIntoHtmlDocument
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert some initial HTML content to demonstrate that the document is HTML based.
        const string initialHtml = "<p>This is a paragraph before the table.</p>";
        builder.InsertHtml(initialHtml, HtmlInsertOptions.PreserveBlocks);

        // Start building a table.
        Table table = builder.StartTable();

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

        // Finish the table.
        builder.EndTable();

        // Optionally set table formatting (e.g., auto‑fit to contents).
        table.AutoFit(AutoFitBehavior.AutoFitToContents);
        table.Alignment = TableAlignment.Center;

        // Save the document as an HTML file.
        doc.Save("Output.html", SaveFormat.Html);
    }
}
