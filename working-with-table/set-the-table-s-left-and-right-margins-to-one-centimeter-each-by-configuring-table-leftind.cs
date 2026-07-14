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

        // Build a simple 2‑cell table.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Left cell");
        builder.InsertCell();
        builder.Write("Right cell");
        builder.EndRow();
        builder.EndTable();

        // Set left indent (margin) to 1 cm (≈28.35 points).
        table.LeftIndent = 28.35;

        // Set right distance (margin) to 1 cm.
        table.DistanceRight = 28.35;

        // Save the document to the local file system.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableMargins.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output file was not created.");

        // The program finishes automatically; no user interaction required.
    }
}
