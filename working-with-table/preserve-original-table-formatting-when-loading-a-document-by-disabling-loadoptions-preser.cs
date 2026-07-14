using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Loading;   // Needed for LoadOptions

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a sample document with a formatted table.
        string originalPath = Path.Combine(artifactsDir, "Original.docx");
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a table and apply cell shading to demonstrate formatting.
        Table table = builder.StartTable();

        // First cell with light blue shading.
        builder.InsertCell();
        builder.CellFormat.Shading.BackgroundPatternColor = Color.LightBlue;
        builder.Write("Cell 1");

        // Second cell with light green shading.
        builder.InsertCell();
        builder.CellFormat.Shading.BackgroundPatternColor = Color.LightGreen;
        builder.Write("Cell 2");

        builder.EndRow();
        builder.EndTable();

        // Save the original document.
        doc.Save(originalPath);

        // Load the document. No special LoadOptions are required to preserve formatting.
        LoadOptions loadOptions = new LoadOptions(); // Default options preserve formatting.
        Document loadedDoc = new Document(originalPath, loadOptions);

        // Verify that the cell shading is still present after loading.
        Table loadedTable = loadedDoc.FirstSection.Body.Tables[0];
        Color firstCellColor = loadedTable.FirstRow.FirstCell.CellFormat.Shading.BackgroundPatternColor;
        if (firstCellColor.ToArgb() != Color.LightBlue.ToArgb())
            throw new Exception("Table formatting was not preserved after loading.");

        // Save the loaded document.
        string loadedPath = Path.Combine(artifactsDir, "Loaded.docx");
        loadedDoc.Save(loadedPath);
    }
}
