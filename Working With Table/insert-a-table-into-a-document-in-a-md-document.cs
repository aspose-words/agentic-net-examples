using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Tables; // Added namespace for Table class

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a heading before the table.
        builder.Writeln("Sample table:");

        // Start a new table.
        Table table = builder.StartTable();

        // ---- First row ----
        builder.InsertCell();               // First cell of the first row.
        builder.Write("Cell 1");

        builder.InsertCell();               // Second cell of the first row.
        builder.Write("Cell 2");

        builder.EndRow();                   // End the first row.

        // ---- Second row ----
        builder.InsertCell();               // First cell of the second row.
        builder.Write("Cell 3");

        builder.InsertCell();               // Second cell of the second row.
        builder.Write("Cell 4");

        builder.EndRow();                   // End the second row.
        builder.EndTable();                 // Finish the table.

        // Configure Markdown save options to export tables as raw HTML (optional).
        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions
        {
            ExportAsHtml = MarkdownExportAsHtml.Tables
        };

        // Save the document as a Markdown file.
        doc.Save("TableInMarkdown.md", saveOptions);
    }
}
