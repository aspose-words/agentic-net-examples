using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new table.
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

        // Apply a built‑in style and auto‑fit the table to its contents.
        table.StyleIdentifier = StyleIdentifier.LightShadingAccent1;
        table.AutoFit(AutoFitBehavior.AutoFitToContents);

        // Save the document as a Word template (.dot).
        doc.Save("TableTemplate.dot");
    }
}
