using System;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Loading; // Required for LoadOptions

public class Program
{
    public static void Main()
    {
        // Create a sample document with a formatted table.
        Document originalDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(originalDoc);

        // Start the table.
        builder.StartTable();

        // First cell with shading and border.
        builder.InsertCell();
        builder.CellFormat.Shading.BackgroundPatternColor = Color.Yellow;
        builder.CellFormat.Borders.LineStyle = LineStyle.Single;
        builder.Writeln("Cell 1");

        // Second cell with different shading.
        builder.InsertCell();
        builder.CellFormat.Shading.BackgroundPatternColor = Color.LightBlue;
        builder.Writeln("Cell 2");

        // End the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the original document.
        const string originalPath = "Original.docx";
        originalDoc.Save(originalPath);

        // Load the document without any special LoadOptions (default preserves formatting).
        LoadOptions loadOptions = new LoadOptions(); // No PreserveFormatting property in this version.
        Document loadedDoc = new Document(originalPath, loadOptions);

        // Verify that the table formatting (shading) is still present.
        Table table = (Table)loadedDoc.GetChild(NodeType.Table, 0, true);
        Cell firstCell = table.Rows[0].Cells[0];
        Color shadingColor = firstCell.CellFormat.Shading.BackgroundPatternColor;

        if (shadingColor.ToArgb() != Color.Yellow.ToArgb())
        {
            throw new InvalidOperationException("Table cell shading was not preserved after loading.");
        }

        // Save the loaded document.
        const string loadedPath = "LoadedPreserved.docx";
        loadedDoc.Save(loadedPath);

        // Validate that both files exist.
        if (!File.Exists(originalPath) || !File.Exists(loadedPath))
        {
            throw new FileNotFoundException("One of the output files was not created.");
        }

        // Indicate successful completion (no interactive output required).
        Console.WriteLine("Document processing completed successfully.");
    }
}
