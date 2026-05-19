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

        // Start a table and add a row with two cells.
        Table table = builder.StartTable();

        // First cell contains Arabic text.
        builder.InsertCell();
        builder.Write("مرحبا بالعالم"); // "Hello World" in Arabic.

        // Second cell contains English text (optional, just for contrast).
        builder.InsertCell();
        builder.Write("Hello World");

        // Finish the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Enable right‑to‑left layout for the table.
        table.Bidi = true;

        // Save the document to the local file system.
        const string fileName = "TableRightToLeft.docx";
        doc.Save(fileName);

        // Verify that the file was created.
        if (!File.Exists(fileName))
            throw new Exception("The document was not saved successfully.");
    }
}
