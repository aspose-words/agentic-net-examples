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

        // Build a simple 2x2 table.
        Table table = builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Write("Cell 3");
        builder.InsertCell();
        builder.Write("Cell 4");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Apply read‑only protection to the entire document.
        doc.Protect(ProtectionType.ReadOnly);

        // Verify that the protection was applied.
        if (doc.ProtectionType != ProtectionType.ReadOnly)
            throw new InvalidOperationException("Document protection was not applied.");

        // Save the protected document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ProtectedTable.docx");
        doc.Save(outputPath);

        // Ensure the file was created successfully.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("Failed to save the protected document.", outputPath);
    }
}
