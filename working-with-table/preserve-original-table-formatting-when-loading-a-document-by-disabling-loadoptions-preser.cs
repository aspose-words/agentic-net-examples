using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Ensure the output directory exists.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // 1. Create a sample document with a formatted table.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table.
        Table table = builder.StartTable();

        // First cell with light blue shading.
        builder.InsertCell();
        builder.CellFormat.Shading.BackgroundPatternColor = Color.LightBlue;
        builder.Write("Cell 1");

        // Second cell with light green shading.
        builder.InsertCell();
        builder.CellFormat.Shading.BackgroundPatternColor = Color.LightGreen;
        builder.Write("Cell 2");

        // End the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the original document.
        string originalPath = Path.Combine(artifactsDir, "Original.docx");
        doc.Save(originalPath);

        // -----------------------------------------------------------------
        // 2. Load the document. No special LoadOptions are required because
        //    the default loading behavior preserves the original table formatting.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(originalPath);

        // Save the loaded document to verify that formatting is preserved.
        string loadedPath = Path.Combine(artifactsDir, "LoadedPreserved.docx");
        loadedDoc.Save(loadedPath);
    }
}
