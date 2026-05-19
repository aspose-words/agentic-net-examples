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

        // Start a table and add a few rows/cells with sample text.
        Table table = builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Write("Cell A1");
        builder.InsertCell();
        builder.Write("Cell A2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Apply a built‑in style (optional, demonstrates style usage).
        table.StyleIdentifier = StyleIdentifier.MediumShading1Accent1;

        // The TableStyleOptions enum does not contain a NoShading member.
        // To disable all cell shading, clear the shading directly.
        table.ClearShading();

        // Save the document to a local file.
        string outputPath = "TableNoShading.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved correctly.");
    }
}
