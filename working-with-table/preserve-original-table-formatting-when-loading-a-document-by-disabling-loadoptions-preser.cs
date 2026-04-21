using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Loading;

public class Program
{
    public static void Main()
    {
        // Define a folder for all generated files.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // 1. Create a sample document that contains a formatted table.
        // -----------------------------------------------------------------
        Document originalDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(originalDoc);

        // Start a table.
        Table table = builder.StartTable();

        // First cell – apply a background shading to demonstrate formatting.
        builder.InsertCell();
        builder.CellFormat.Shading.BackgroundPatternColor = System.Drawing.Color.LightBlue;
        builder.Write("First cell");

        // Second cell – different shading.
        builder.InsertCell();
        builder.CellFormat.Shading.BackgroundPatternColor = System.Drawing.Color.LightGreen;
        builder.Write("Second cell");

        // Finish the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the original document.
        string originalPath = Path.Combine(artifactsDir, "Original.docx");
        originalDoc.Save(originalPath);

        // -----------------------------------------------------------------
        // 2. Load the document. No PreserveFormatting property exists,
        //    so we simply load with default LoadOptions.
        // -----------------------------------------------------------------
        LoadOptions loadOptions = new LoadOptions(); // default options

        Document loadedDoc = new Document(originalPath, loadOptions);

        // Optional verification: ensure the first cell still has the LightBlue shading.
        Cell firstCell = (Cell)loadedDoc.GetChild(NodeType.Cell, 0, true);
        System.Drawing.Color shadingColor = firstCell.CellFormat.Shading.BackgroundPatternColor;
        if (shadingColor.ToArgb() != System.Drawing.Color.LightBlue.ToArgb())
            throw new Exception("Table formatting was not preserved after loading.");

        // Save the loaded document.
        string loadedPath = Path.Combine(artifactsDir, "Loaded.docx");
        loadedDoc.Save(loadedPath);

        // -----------------------------------------------------------------
        // 3. Validate that the output file exists.
        // -----------------------------------------------------------------
        if (!File.Exists(loadedPath))
            throw new FileNotFoundException("The loaded document was not saved correctly.", loadedPath);
    }
}
