using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Saving;

class JoinTablesInTxt
{
    static void Main()
    {
        // Load the source document that already contains two tables.
        // (Replace with the actual path to your document.)
        Document doc = new Document("SourceDocument.docx");

        // Get the collection of tables in the first section.
        TableCollection tables = doc.FirstSection.Body.Tables;

        // Ensure there are at least two tables to join.
        if (tables.Count < 2)
        {
            Console.WriteLine("The document must contain at least two tables.");
            return;
        }

        // Reference to the first and second tables.
        Table firstTable = tables[0];
        Table secondTable = tables[1];

        // Move all rows from the second table to the first table.
        // Import each row into the document to preserve formatting.
        foreach (Row row in secondTable.Rows)
        {
            // Import the row (deep clone) into the current document.
            Row importedRow = (Row)doc.ImportNode(row, true);
            firstTable.Rows.Add(importedRow);
        }

        // Remove the now-empty second table from the document.
        secondTable.Remove();

        // Save the resulting document as plain text.
        // The tables will be rendered as plain text tables in the output file.
        TxtSaveOptions saveOptions = new TxtSaveOptions
        {
            // Export headers/footers as needed; here we keep the default (None).
            ExportHeadersFootersMode = TxtExportHeadersFootersMode.None
        };

        doc.Save("JoinedTables.txt", saveOptions);
    }
}
