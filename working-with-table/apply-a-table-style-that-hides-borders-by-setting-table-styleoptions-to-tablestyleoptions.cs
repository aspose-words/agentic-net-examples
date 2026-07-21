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

        // First row.
        builder.InsertCell();
        builder.Write("Cell 1,1");
        builder.InsertCell();
        builder.Write("Cell 1,2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Write("Cell 2,1");
        builder.InsertCell();
        builder.Write("Cell 2,2");
        builder.EndTable(); // Ends the table and returns the Table object.

        // Apply a style (any style identifier) – here we use a built‑in style.
        table.StyleIdentifier = StyleIdentifier.LightShadingAccent1;

        // The TableStyleOptions enum does not contain a NoBorders member.
        // To hide borders we clear them directly.
        table.ClearBorders();

        // Optionally set StyleOptions to None (no conditional formatting).
        table.StyleOptions = TableStyleOptions.None;

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableNoBorders.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved correctly.");
    }
}
