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

        // Insert a cell containing Arabic text.
        builder.InsertCell();
        builder.Write("مرحبا بالعالم"); // Arabic: "Hello World"

        // End the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Enable right‑to‑left layout for the table.
        table.Bidi = true;

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableRightToLeft.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("Failed to create the output document.");

        // Reload the document and confirm the table direction.
        Document loadedDoc = new Document(outputPath);
        Table loadedTable = loadedDoc.GetChildNodes(NodeType.Table, true)[0] as Table;
        if (loadedTable == null || !loadedTable.Bidi)
            throw new Exception("Table direction was not set to right‑to‑left.");
    }
}
