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

        // Move the builder cursor to the end of the document.
        builder.MoveToDocumentEnd();

        // Start a new table at the current cursor position.
        Table table = builder.StartTable();

        // ---- First row ----
        builder.InsertCell();                     // First cell of the first row.
        builder.Write("Cell 1, Row 1");
        builder.InsertCell();                     // Second cell of the first row.
        builder.Write("Cell 2, Row 1");
        builder.EndRow();                         // End the first row.

        // ---- Second row ----
        builder.InsertCell();                     // First cell of the second row.
        builder.Write("Cell 1, Row 2");
        builder.InsertCell();                     // Second cell of the second row.
        builder.Write("Cell 2, Row 2");
        builder.EndRow();                         // End the second row.

        // Finish the table.
        builder.EndTable();

        // Save the document to a file.
        string outputPath = "TableAtEnd.docx";
        doc.Save(outputPath);

        // Simple validation to ensure the file was created.
        if (!File.Exists(outputPath))
            throw new Exception($"Failed to create the output file: {outputPath}");
    }
}
