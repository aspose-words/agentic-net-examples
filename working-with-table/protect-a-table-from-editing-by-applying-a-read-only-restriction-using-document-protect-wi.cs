using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to construct a simple table.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.StartTable();

        // First row, first cell.
        builder.InsertCell();
        builder.Writeln("Header 1");

        // First row, second cell.
        builder.InsertCell();
        builder.Writeln("Header 2");
        builder.EndRow();

        // Second row, first cell.
        builder.InsertCell();
        builder.Writeln("Cell 1");

        // Second row, second cell.
        builder.InsertCell();
        builder.Writeln("Cell 2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Apply read‑only protection to the whole document (which includes the table).
        // The second parameter is the password; an empty string means no password.
        doc.Protect(ProtectionType.ReadOnly, "myPassword");

        // Define output file path.
        string outputPath = "ProtectedTable.docx";

        // Save the protected document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception($"Failed to create the output file: {outputPath}");
        }

        // Load the saved document to confirm protection type.
        Document loadedDoc = new Document(outputPath);
        if (loadedDoc.ProtectionType != ProtectionType.ReadOnly)
        {
            throw new Exception("The document is not protected as expected.");
        }

        // Program completed successfully.
    }
}
