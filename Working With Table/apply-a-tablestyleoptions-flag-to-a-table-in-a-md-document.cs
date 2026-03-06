using System;
using Aspose.Words;
using Aspose.Words.Tables;

class ApplyTableStyleOptions
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start building a table.
        Table table = builder.StartTable();

        // Insert the first row and cells (required before setting table formatting).
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.EndRow();

        // Insert a data row.
        builder.InsertCell();
        builder.Write("Data 1");
        builder.InsertCell();
        builder.Write("Data 2");
        builder.EndRow();

        // End the table construction.
        builder.EndTable();

        // Apply a built‑in style identifier (optional, but demonstrates style usage).
        table.StyleIdentifier = StyleIdentifier.MediumShading1Accent1;

        // Apply the desired TableStyleOptions flags.
        // Example: apply formatting to the first row, first column, and row banding.
        table.StyleOptions = TableStyleOptions.FirstRow |
                              TableStyleOptions.FirstColumn |
                              TableStyleOptions.RowBands;

        // Optionally auto‑fit the table to its contents.
        table.AutoFit(AutoFitBehavior.AutoFitToContents);

        // Save the document. Adjust the path as needed.
        string outputPath = "TableWithStyleOptions.docx";
        doc.Save(outputPath);
    }
}
