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

        // Start building a table.
        builder.StartTable();

        // Insert a cell and add some text.
        builder.InsertCell();
        builder.Writeln("Cell with light gray shading");

        // Retrieve the current cell that was just created.
        Cell cell = builder.CurrentParagraph.ParentNode as Cell;
        if (cell == null)
            throw new InvalidOperationException("Current node is not a table cell.");

        // Apply light gray shading to the cell.
        cell.CellFormat.Shading.ForegroundPatternColor = Color.LightGray;
        cell.CellFormat.Shading.BackgroundPatternColor = Color.Empty;
        cell.CellFormat.Shading.Texture = TextureIndex.TextureNone; // Solid fill.

        // End the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document to disk.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Verify that the file was created successfully.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");
    }
}
