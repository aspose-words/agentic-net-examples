using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class SetTableBidirectional
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
        builder.EndTable();

        // Set the table to be right‑to‑left (bidirectional).
        table.Bidi = true;

        // Define output path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "BidirectionalTable.docx");

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output file was not created.");

        // Optional: inform that the process completed successfully.
        Console.WriteLine("Document saved to: " + outputPath);
    }
}
