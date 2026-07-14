using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing; // For TextureIndex
using System.Drawing; // For Color

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 1x2 table.
        Table table = builder.StartTable();

        // First cell.
        builder.InsertCell();
        builder.Write("Cell 1");

        // Second cell.
        builder.InsertCell();
        builder.Write("Cell 2");

        // Finish the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Apply light gray shading to the first cell.
        Cell firstCell = table.FirstRow.FirstCell;
        firstCell.CellFormat.Shading.Texture = TextureIndex.TextureSolid;
        firstCell.CellFormat.Shading.BackgroundPatternColor = Color.LightGray;

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "CellShading.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("The output file was not created.");
        }
    }
}
