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

        // Start a table.
        builder.StartTable();

        // First cell – Arabic text with right‑to‑left direction.
        builder.InsertCell();
        // Enable right‑to‑left paragraph direction.
        builder.ParagraphFormat.Bidi = true;
        builder.Write("مرحبا بالعالم"); // Arabic greeting.

        // Second cell – regular left‑to‑right text.
        builder.InsertCell();
        // Reset paragraph direction to default (left‑to‑right).
        builder.ParagraphFormat.Bidi = false;
        builder.Write("Hello world!");

        // Finish the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document to a file in the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "CellTextDirection.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved correctly.");
    }
}
