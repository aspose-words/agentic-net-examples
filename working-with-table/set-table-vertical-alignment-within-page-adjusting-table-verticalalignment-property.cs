using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 1‑row, 1‑cell table.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Centered table on the page");
        builder.EndRow();
        builder.EndTable();

        // Align the floating table vertically to the center of the page.
        table.RelativeVerticalAlignment = VerticalAlignment.Center;

        // (Optional) Align the table horizontally to the center as well.
        table.RelativeHorizontalAlignment = HorizontalAlignment.Center;

        // Save the result.
        doc.Save("TableVerticalAlignment.docx");
    }
}
