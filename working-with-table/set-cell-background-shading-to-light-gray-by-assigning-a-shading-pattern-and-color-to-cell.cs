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

        // Start a new table.
        builder.StartTable();

        // Insert first cell.
        builder.InsertCell();
        // Set the cell background shading to light gray.
        Cell cell = builder.CurrentParagraph.ParentNode as Cell;
        if (cell == null)
            throw new InvalidOperationException("Current node is not a cell.");

        cell.CellFormat.Shading.Texture = TextureIndex.TextureSolid;               // Solid fill pattern.
        cell.CellFormat.Shading.ForegroundPatternColor = Color.LightGray;        // Light gray color.

        // Add some text to the cell.
        builder.Writeln("Shaded Cell");

        // End the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Define output path.
        string outputPath = "TableShading.docx";

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output file was not created.", outputPath);
    }
}
