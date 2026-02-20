using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Tables;

class JoinTablesInMarkdown
{
    static void Main()
    {
        // Paths to the input and output files.
        string inputPath = @"C:\Docs\InputDocument.docx";
        string outputPath = @"C:\Docs\JoinedTables.md";

        // Load the existing document.
        Document doc = new Document(inputPath);

        // Ensure the document has at least two tables.
        if (doc.FirstSection?.Body?.Tables?.Count < 2)
        {
            Console.WriteLine("The document does not contain two tables to join.");
            return;
        }

        // Retrieve the first two tables.
        Table firstTable = doc.FirstSection.Body.Tables[0];
        Table secondTable = doc.FirstSection.Body.Tables[1];

        // Append all rows from the second table to the first table.
        // Clone each row to preserve formatting and content.
        foreach (Row row in secondTable.Rows)
        {
            // Clone the row (deep clone) and add it to the first table.
            firstTable.Rows.Add(row.Clone(true));
        }

        // Remove the now redundant second table from the document.
        secondTable.Remove();

        // Save the document as Markdown, exporting tables as raw HTML.
        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions
        {
            ExportAsHtml = MarkdownExportAsHtml.Tables
        };
        doc.Save(outputPath, saveOptions);

        Console.WriteLine("Tables have been joined and saved to Markdown.");
    }
}
