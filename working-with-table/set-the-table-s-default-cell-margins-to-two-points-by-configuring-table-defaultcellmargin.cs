using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 2x2 table.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1,1");
        builder.InsertCell();
        builder.Write("Cell 1,2");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("Cell 2,1");
        builder.InsertCell();
        builder.Write("Cell 2,2");
        builder.EndRow();
        builder.EndTable();

        // Set the default cell margins (padding) for the table to 2 points on each side.
        // These properties affect all cells that do not have an explicit padding set.
        table.LeftPadding = 2.0;
        table.RightPadding = 2.0;
        table.TopPadding = 2.0;
        table.BottomPadding = 2.0;

        // Define output path.
        string artifactsDir = Path.Combine(Environment.CurrentDirectory, "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string outPath = Path.Combine(artifactsDir, "TableWithDefaultCellMargins.docx");

        // Save the document.
        doc.Save(outPath);

        // Verify that the file was created.
        if (!File.Exists(outPath))
            throw new InvalidOperationException("The output document was not saved correctly.");
    }
}
