using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using System.Drawing;

public class PreserveTableFormattingExample
{
    public static void Main()
    {
        // Define a folder for all generated files.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Paths for the original and the re‑loaded documents.
        string originalPath = Path.Combine(artifactsDir, "Original.docx");
        string reloadedPath = Path.Combine(artifactsDir, "ReloadedPreserved.docx");

        // -------------------------------------------------
        // 1. Create a sample document that contains a table.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a 1‑row, 2‑cell table with distinct cell shading.
        builder.StartTable();

        builder.InsertCell();
        builder.CellFormat.Shading.BackgroundPatternColor = Color.LightBlue;
        builder.Write("Cell 1");

        builder.InsertCell();
        builder.CellFormat.Shading.BackgroundPatternColor = Color.LightGreen;
        builder.Write("Cell 2");

        builder.EndRow();
        builder.EndTable();

        // Save the original document.
        doc.Save(originalPath);

        // -------------------------------------------------
        // 2. Load the document.
        // -------------------------------------------------
        // In this version of Aspose.Words the PreserveFormatting option is enabled by default,
        // so loading without explicit LoadOptions preserves the original table formatting.
        Document loadedDoc = new Document(originalPath);

        // -------------------------------------------------
        // 3. Verify that the table formatting is still present.
        // -------------------------------------------------
        Table loadedTable = loadedDoc.FirstSection.Body.Tables[0];
        Cell firstCell = loadedTable.FirstRow.FirstCell;
        Cell secondCell = loadedTable.FirstRow.LastCell;

        // The background colors should match the ones we set earlier.
        if (firstCell.CellFormat.Shading.BackgroundPatternColor.ToArgb() != Color.LightBlue.ToArgb() ||
            secondCell.CellFormat.Shading.BackgroundPatternColor.ToArgb() != Color.LightGreen.ToArgb())
        {
            throw new InvalidOperationException("Table formatting was not preserved after loading.");
        }

        // -------------------------------------------------
        // 4. Save the re‑loaded document.
        // -------------------------------------------------
        loadedDoc.Save(reloadedPath);

        // Simple validation that the output file exists.
        if (!File.Exists(reloadedPath))
            throw new FileNotFoundException("The reloaded document was not saved correctly.", reloadedPath);
    }
}
