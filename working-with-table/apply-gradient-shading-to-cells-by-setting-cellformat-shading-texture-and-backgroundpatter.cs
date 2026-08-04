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

        // Start a table.
        Table table = builder.StartTable();

        // ---- First row ----
        // First cell with diagonal cross texture and light blue background.
        builder.InsertCell();
        builder.CellFormat.Shading.Texture = TextureIndex.TextureDiagonalCross;
        builder.CellFormat.Shading.BackgroundPatternColor = Color.LightBlue;
        builder.Writeln("Cell 1");

        // Second cell with horizontal texture and light green background.
        builder.InsertCell();
        builder.CellFormat.Shading.Texture = TextureIndex.TextureHorizontal;
        builder.CellFormat.Shading.BackgroundPatternColor = Color.LightGreen;
        builder.Writeln("Cell 2");

        // End the first row.
        builder.EndRow();

        // ---- Second row ----
        // First cell with vertical texture and light coral background.
        builder.InsertCell();
        builder.CellFormat.Shading.Texture = TextureIndex.TextureVertical;
        builder.CellFormat.Shading.BackgroundPatternColor = Color.LightCoral;
        builder.Writeln("Cell 3");

        // Second cell with cross texture and light yellow background.
        builder.InsertCell();
        builder.CellFormat.Shading.Texture = TextureIndex.TextureCross;
        builder.CellFormat.Shading.BackgroundPatternColor = Color.LightYellow;
        builder.Writeln("Cell 4");

        // End the second row.
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "GradientShadingTable.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("Document was not saved successfully.");
    }
}
