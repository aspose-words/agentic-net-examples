using System;
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

        // First row with two cells.
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();

        // Second row with two cells.
        builder.InsertCell();
        builder.Write("Cell 3");
        builder.InsertCell();
        builder.Write("Cell 4");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Apply a built‑in style to the table (optional, demonstrates StyleIdentifier usage).
        table.StyleIdentifier = StyleIdentifier.MediumShading1Accent1;

        // Disable all cell shading. The TableStyleOptions enum does not contain a NoShading member,
        // so we clear shading directly on the table.
        table.ClearShading();

        // Save the document to the local file system.
        const string outputPath = "Table_NoShading.docx";
        doc.Save(outputPath);
    }
}
