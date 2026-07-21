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
        Table table = builder.StartTable();

        // Insert the first cell.
        builder.InsertCell();

        // Set paragraph direction to right‑to‑left for Arabic content.
        builder.ParagraphFormat.Bidi = true;

        // Write Arabic text.
        builder.Write("مرحبا بالعالم"); // "Hello World" in Arabic.

        // End the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Define output path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ArabicTable.docx");

        // Save the document.
        doc.Save(outputPath);

        // Simple verification that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output file was not created.");
    }
}
