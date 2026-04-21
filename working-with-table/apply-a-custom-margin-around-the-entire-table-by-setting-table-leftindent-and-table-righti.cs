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

        // Start building a table.
        Table table = builder.StartTable();

        // First row, first cell.
        builder.InsertCell();
        builder.Write("Cell 1");

        // First row, second cell.
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Apply a custom left margin (indent) to the whole table.
        table.LeftIndent = 30; // points

        // Note: Table.RightIndent is not available in this API version and must not be used.

        // Prepare output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string outputPath = Path.Combine(artifactsDir, "TableWithMargin.docx");

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("Failed to create the output document.");

        // Reload the document to confirm the left indent was applied.
        Document loadedDoc = new Document(outputPath);
        Table loadedTable = loadedDoc.FirstSection.Body.Tables[0];
        if (Math.Abs(loadedTable.LeftIndent - 30) > 0.01)
            throw new Exception("LeftIndent was not applied correctly.");
    }
}
