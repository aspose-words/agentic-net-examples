using Aspose.Words;
using Aspose.Words.Tables;
using System.Drawing;

class Program
{
    static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 2x2 table.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();
        builder.InsertCell();
        builder.Write("Cell 3");
        builder.InsertCell();
        builder.Write("Cell 4");
        builder.EndTable();

        // Apply background shading to the entire table.
        table.ClearShading(); // Remove any existing shading.
        table.SetShading(TextureIndex.TextureSolid, Color.Empty, Color.LightBlue); // Set solid background color.

        // Save the document.
        doc.Save("TableShading.docx");
    }
}
