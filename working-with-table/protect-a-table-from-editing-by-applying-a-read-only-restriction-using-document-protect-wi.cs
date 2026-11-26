using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 2x2 table.
        builder.StartTable();

        // First row, first cell.
        builder.InsertCell();
        builder.Writeln("Cell 1");
        builder.EndRow();

        // Second row, first cell.
        builder.InsertCell();
        builder.Writeln("Cell 2");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Apply read‑only protection to the document (the table becomes read‑only as well).
        doc.Protect(ProtectionType.ReadOnly, "myPassword");

        // Save the protected document.
        string outputPath = "ProtectedTable.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The protected document was not saved.");

        // Indicate successful completion.
        Console.WriteLine("Document saved and protected successfully.");
    }
}
