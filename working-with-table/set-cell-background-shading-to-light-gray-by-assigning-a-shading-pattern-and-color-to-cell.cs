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

        // Build a simple table with one cell.
        builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Cell with light gray shading");

        // Retrieve the cell that was just created.
        Cell cell = builder.CurrentParagraph.ParentNode as Cell;
        if (cell == null)
            throw new InvalidOperationException("Current node is not a table cell.");

        // Apply solid shading with a light gray color.
        cell.CellFormat.Shading.Texture = TextureIndex.TextureNone; // solid fill
        cell.CellFormat.Shading.ForegroundPatternColor = Color.LightGray;

        // Finish the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document to disk.
        string outputPath = "TableShading.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output file was not created.", outputPath);
    }
}
