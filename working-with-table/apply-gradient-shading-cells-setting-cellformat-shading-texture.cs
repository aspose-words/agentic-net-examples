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

        // Start a table.
        builder.StartTable();

        // Apply gradient-like shading to the first cell.
        builder.CellFormat.Shading.Texture = TextureIndex.TextureDiagonalCross;
        builder.CellFormat.Shading.BackgroundPatternColor = Color.LightCoral;

        // Insert the first cell with the shading applied.
        builder.InsertCell();
        builder.Writeln("Cell 1");

        // Change shading for the second cell.
        builder.CellFormat.Shading.Texture = TextureIndex.TextureHorizontal;
        builder.CellFormat.Shading.BackgroundPatternColor = Color.LightBlue;

        // Insert the second cell.
        builder.InsertCell();
        builder.Writeln("Cell 2");

        // End the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document.
        doc.Save("GradientShadingTable.docx");
    }
}
