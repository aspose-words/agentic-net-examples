using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple table with one cell.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Sample cell");
        builder.EndTable();

        // Set the distance between the table top and surrounding text to 3 points.
        // This property controls the vertical placement of the table relative to surrounding content.
        table.DistanceTop = 3.0;

        // Save the document to the desired location.
        string outputPath = Path.Combine("Artifacts", "TableTopMargin.docx");
        doc.Save(outputPath);
    }
}
