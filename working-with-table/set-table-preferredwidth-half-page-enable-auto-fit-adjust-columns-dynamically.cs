using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new document and a DocumentBuilder to construct its contents.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple table with two cells.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();
        builder.EndTable();

        // Set the table's preferred width to 50 % of the page width.
        table.PreferredWidth = PreferredWidth.FromPercent(50);

        // Enable auto‑fit so the columns can adjust dynamically to their contents.
        table.AllowAutoFit = true;
        table.AutoFit(AutoFitBehavior.AutoFitToContents);

        // Save the document.
        doc.Save("TablePreferredWidthAutoFit.docx");
    }
}
