using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a sample document with a formatted table.
        Document originalDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(originalDoc);

        // Start the table.
        builder.StartTable();

        // First cell with shading.
        builder.InsertCell();
        builder.CellFormat.Shading.BackgroundPatternColor = Color.Yellow;
        builder.Writeln("Cell 1");

        // Second cell without special formatting.
        builder.InsertCell();
        builder.Writeln("Cell 2");

        // End the first row.
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Writeln("Cell 3");
        builder.InsertCell();
        builder.Writeln("Cell 4");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Save the original document.
        string originalPath = "Original.docx";
        originalDoc.Save(originalPath);

        // Load the document without any special LoadOptions.
        // The default loading behavior preserves the original formatting.
        Document loadedDoc = new Document(originalPath);

        // Verify that the shading of the first cell is still present.
        Table table = loadedDoc.GetChild(NodeType.Table, 0, true) as Table;
        if (table == null)
            throw new Exception("Table was not loaded.");

        Cell firstCell = table.Rows[0].Cells[0];
        Color shadingColor = firstCell.CellFormat.Shading.BackgroundPatternColor;

        if (shadingColor.ToArgb() != Color.Yellow.ToArgb())
            throw new Exception("Table formatting was not preserved after loading.");

        // Save the loaded document.
        string loadedPath = "Loaded.docx";
        loadedDoc.Save(loadedPath);

        // Validate that the output files exist.
        if (!File.Exists(originalPath) || !File.Exists(loadedPath))
            throw new Exception("Output files were not created successfully.");
    }
}
