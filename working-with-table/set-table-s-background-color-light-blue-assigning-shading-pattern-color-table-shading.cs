using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new blank document.
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

        // Remove any existing shading from the table.
        table.ClearShading();

        // Apply a solid light‑blue background to the whole table.
        // TextureSolid creates a solid fill; the foreground color is ignored, background color sets the fill.
        table.SetShading(TextureIndex.TextureSolid, Color.Empty, Color.LightBlue);

        // Save the document.
        doc.Save("TableWithLightBlueBackground.docx");
    }
}
