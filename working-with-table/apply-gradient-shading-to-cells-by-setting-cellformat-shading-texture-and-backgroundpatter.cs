using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table.
        Table table = builder.StartTable();

        // First row - apply gradient-like shading to each cell.
        builder.InsertCell();
        builder.CellFormat.Shading.Texture = TextureIndex.TextureDiagonalCross;
        builder.CellFormat.Shading.BackgroundPatternColor = Color.LightBlue;
        builder.Write("Cell 1");

        builder.InsertCell();
        builder.CellFormat.Shading.Texture = TextureIndex.TextureDiagonalCross;
        builder.CellFormat.Shading.BackgroundPatternColor = Color.LightCoral;
        builder.Write("Cell 2");

        builder.EndRow();

        // Second row - normal cells without special shading.
        builder.InsertCell();
        builder.Write("Cell 3");

        builder.InsertCell();
        builder.Write("Cell 4");

        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Save the document to a file in the current directory.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "GradientShadingTable.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The document was not saved successfully.");
    }
}
