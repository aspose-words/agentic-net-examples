using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;   // Needed for the Table class

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Build a simple 2x2 table using DocumentBuilder.
        DocumentBuilder builder = new DocumentBuilder(doc);
        Table table = builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Write("Cell 1,1");
        builder.InsertCell();
        builder.Write("Cell 1,2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Write("Cell 2,1");
        builder.InsertCell();
        builder.Write("Cell 2,2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Apply read‑only protection to the document (which includes the table).
        doc.Protect(ProtectionType.ReadOnly);

        // Save the protected document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ProtectedTable.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Failed to create the output document.");

        // Load the saved document and confirm the protection type.
        Document loaded = new Document(outputPath);
        if (loaded.ProtectionType != ProtectionType.ReadOnly)
            throw new InvalidOperationException("Document protection was not applied correctly.");
    }
}
