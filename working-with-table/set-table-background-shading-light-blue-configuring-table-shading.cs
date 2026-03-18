using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new document and a builder to construct content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 1x1 table.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Sample cell");
        builder.EndRow();
        builder.EndTable();

        // Apply solid light‑blue background shading to the whole table.
        // TextureSolid indicates a solid fill; foreground color is not used, so we pass Color.Empty.
        table.SetShading(TextureIndex.TextureSolid, Color.Empty, Color.LightBlue);

        // Save the document.
        doc.Save("TableWithLightBlueShading.docx");
    }
}
