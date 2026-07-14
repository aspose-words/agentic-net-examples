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

        // Add some initial text so the document is not empty.
        builder.Writeln("Document before the table.");

        // Move the builder cursor to the end of the document.
        builder.MoveToDocumentEnd();

        // Start a table at the current cursor position (document end).
        Table table = builder.StartTable();

        // First row with two cells.
        builder.InsertCell();
        builder.Write("Cell 1, Row 1");
        builder.InsertCell();
        builder.Write("Cell 2, Row 1");
        builder.EndRow();

        // Second row with two cells.
        builder.InsertCell();
        builder.Write("Cell 1, Row 2");
        builder.InsertCell();
        builder.Write("Cell 2, Row 2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document to a file in the current working directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "OutputTable.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The document was not saved correctly.");

        // Output the location of the saved file.
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
