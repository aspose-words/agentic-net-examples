using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the existing MHTML document.
        Document doc = new Document("input.mhtml");

        // Create a DocumentBuilder to edit the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Position the builder where the table should be inserted.
        builder.MoveToDocumentEnd();

        // Start a new table.
        Table table = builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Optional: set title and description for accessibility.
        table.Title = "Sample Table";
        table.Description = "Demonstrates inserting a table into an MHTML document";

        // Save the modified document back to MHTML.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
        {
            // Export table widths as relative values to keep the layout flexible.
            TableWidthOutputMode = HtmlElementSizeOutputMode.RelativeOnly
        };
        doc.Save("output.mhtml", saveOptions);
    }
}
