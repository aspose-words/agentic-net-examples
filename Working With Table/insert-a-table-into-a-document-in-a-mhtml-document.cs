using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Use DocumentBuilder to simplify node insertion.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new table.
        Table table = builder.StartTable();

        // ---- First row ----
        builder.InsertCell();                     // First cell of the first row.
        builder.Write("Cell 1,1");                // Add text to the cell.

        builder.InsertCell();                     // Second cell of the first row.
        builder.Write("Cell 1,2");
        builder.EndRow();                         // End the first row.

        // ---- Second row ----
        builder.InsertCell();                     // First cell of the second row.
        builder.Write("Cell 2,1");

        builder.InsertCell();                     // Second cell of the second row.
        builder.Write("Cell 2,2");
        builder.EndRow();                         // End the second row.

        // Finish the table.
        builder.EndTable();

        // Optional: set title and description for accessibility.
        table.Title = "Sample Table";
        table.Description = "Demonstrates inserting a table into an MHTML document.";

        // Save the document as MHTML.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
        {
            // Export table widths as relative percentages to keep layout flexible.
            TableWidthOutputMode = HtmlElementSizeOutputMode.RelativeOnly
        };

        doc.Save("TableInMhtml.mht", saveOptions);
    }
}
