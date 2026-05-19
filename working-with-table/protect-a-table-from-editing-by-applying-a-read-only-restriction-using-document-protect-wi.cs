using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;   // Needed for Table class

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a simple table with two cells.
        builder.Writeln("Sample table:");
        Table table = builder.StartTable();   // Start the table and keep a reference.
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();
        builder.EndTable();

        // Apply read‑only protection to the whole document (including the table).
        doc.Protect(ProtectionType.ReadOnly);

        // Define the output path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ProtectedTable.docx");

        // Save the protected document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The document was not saved correctly.");
    }
}
