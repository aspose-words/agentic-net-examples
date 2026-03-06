using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Create a DocumentBuilder which will be used to insert content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new table.
        builder.StartTable();

        // First cell of the first row.
        builder.InsertCell();
        builder.Write("Cell 1");

        // Second cell of the first row.
        builder.InsertCell();
        builder.Write("Cell 2");

        // End the first row.
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Save the document as a Markdown file.
        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions();
        doc.Save("TableInMarkdown.md", saveOptions);
    }
}
