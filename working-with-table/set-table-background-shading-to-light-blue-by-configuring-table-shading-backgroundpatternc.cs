using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 2x1 table.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("First cell");
        builder.InsertCell();
        builder.Write("Second cell");
        builder.EndRow();
        builder.EndTable();

        // Set the whole table's background shading to light blue.
        // Use Table.SetShading since Table.Shading property does not exist.
        table.SetShading(TextureIndex.TextureSolid, Color.Empty, Color.LightBlue);

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableShading.docx");
        doc.Save(outputPath);
    }
}
