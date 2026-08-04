using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 1‑row, 2‑cell table.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();
        builder.EndTable();

        // Set the top (distance to preceding text) and bottom (distance to following text) spacing of the table in points.
        table.DistanceTop = 12;    // 12 points = 1/6 inch
        table.DistanceBottom = 24; // 24 points = 1/3 inch

        // Save the document.
        doc.Save("TableSpacing.docx");
    }
}
