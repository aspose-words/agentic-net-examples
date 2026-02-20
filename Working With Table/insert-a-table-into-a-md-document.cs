using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Attach a DocumentBuilder to the document for building content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a heading before the table.
        builder.Writeln("Sample table:");

        // Start a new table.
        Table table = builder.StartTable();

        // ---- First row ----
        builder.InsertCell();                                   // First cell.
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Right;
        builder.Write("Cell1");                                 // Write content.
        builder.InsertCell();                                   // Second cell.
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        builder.Write("Cell2");
        builder.EndRow();                                       // End first row.

        // ---- Second row ----
        builder.InsertCell();
        builder.Write("Data1");
        builder.InsertCell();
        builder.Write("Data2");
        builder.EndRow();                                       // End second row.

        // Finish the table.
        builder.EndTable();

        // Prepare Markdown save options.
        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions();

        // Optionally export tables as raw HTML within the Markdown.
        // Uncomment the following line if raw HTML tables are desired:
        // saveOptions.ExportAsHtml = MarkdownExportAsHtml.Tables;

        // Save the document as a Markdown file.
        doc.Save("TableDocument.md", saveOptions);
    }
}
