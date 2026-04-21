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

        // Start a table and add a couple of cells with sample text.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("English text");
        builder.InsertCell();
        builder.Write("עברית"); // Hebrew text
        builder.EndRow();
        builder.EndTable();

        // Set the table to right‑to‑left layout.
        table.Bidi = true;

        // Define output path and ensure the directory exists.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "TableBidi.docx");

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output file was not created.");

        // Optionally, you could load the document again and check the Bidi flag.
        Document loaded = new Document(outputPath);
        Table loadedTable = loaded.FirstSection.Body.Tables[0];
        if (!loadedTable.Bidi)
            throw new InvalidOperationException("The table Bidi property was not set correctly.");
    }
}
