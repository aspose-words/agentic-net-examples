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

        // First cell with Arabic text.
        builder.InsertCell();
        builder.Write("مرحبا"); // "Hello"

        // Second cell with Arabic text.
        builder.InsertCell();
        builder.Write("كيف حالك؟"); // "How are you?"

        // End the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Set the table direction to right‑to‑left.
        table.Bidi = true;

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "RightToLeftTable.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not saved.");

        // Indicate successful completion.
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
