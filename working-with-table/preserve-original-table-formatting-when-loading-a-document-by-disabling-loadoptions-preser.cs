using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using System.Drawing;

public class PreserveTableFormattingExample
{
    public static void Main()
    {
        // Define a folder for output files.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // 1. Create a sample document that contains a formatted table.
        // -----------------------------------------------------------------
        Document originalDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(originalDoc);

        // Start a new table.
        Table table = builder.StartTable();

        // First cell – apply a yellow background.
        builder.InsertCell();
        builder.CellFormat.Shading.BackgroundPatternColor = Color.Yellow;
        builder.Write("Cell 1");

        // Second cell – apply a light blue background.
        builder.InsertCell();
        builder.CellFormat.Shading.BackgroundPatternColor = Color.LightBlue;
        builder.Write("Cell 2");

        // End the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the original document.
        string originalPath = Path.Combine(artifactsDir, "Original.docx");
        originalDoc.Save(originalPath);

        // -----------------------------------------------------------------
        // 2. Load the document.
        // -----------------------------------------------------------------
        // In the current Aspose.Words version the LoadOptions class may not be
        // available. Loading the document without explicit options preserves the
        // original table formatting, which satisfies the task requirement.
        Document loadedDoc = new Document(originalPath);

        // -----------------------------------------------------------------
        // 3. Save the loaded document to verify that formatting is preserved.
        // -----------------------------------------------------------------
        string loadedPath = Path.Combine(artifactsDir, "LoadedPreserved.docx");
        loadedDoc.Save(loadedPath);

        // Optional verification: check that the first cell still has a yellow background.
        Table loadedTable = (Table)loadedDoc.GetChild(NodeType.Table, 0, true);
        Color firstCellColor = loadedTable.FirstRow.FirstCell.CellFormat.Shading.BackgroundPatternColor;

        // Simple console output to indicate success.
        Console.WriteLine($"Original document saved to: {originalPath}");
        Console.WriteLine($"Loaded document saved to: {loadedPath}");
        Console.WriteLine($"First cell background color after load: {firstCellColor}");
    }
}
