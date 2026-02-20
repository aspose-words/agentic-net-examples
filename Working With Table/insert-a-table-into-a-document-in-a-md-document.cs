using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Tables;

class InsertTableIntoMarkdown
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a heading before the table.
        builder.Writeln("Sample table:");

        // Start a new table.
        builder.StartTable();

        // First cell – right aligned text.
        builder.InsertCell();
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Right;
        builder.Write("Cell1");

        // Second cell – center aligned text.
        builder.InsertCell();
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        builder.Write("Cell2");

        // End the first (and only) row.
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Configure Markdown save options to export tables as raw HTML.
        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions
        {
            ExportAsHtml = MarkdownExportAsHtml.Tables
        };

        // Save the document as a Markdown file.
        doc.Save("SampleTable.md", saveOptions);
    }
}
