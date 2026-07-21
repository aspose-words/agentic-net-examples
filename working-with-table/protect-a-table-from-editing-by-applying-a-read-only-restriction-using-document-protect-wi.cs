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

        // Apply read‑only protection to the whole document.
        // The table will be non‑editable in Microsoft Word.
        doc.Protect(ProtectionType.ReadOnly);

        // Define the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ProtectedTable.docx");

        // Save the protected document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not created.");

        // Load the document again to confirm the protection type.
        Document loadedDoc = new Document(outputPath);
        if (loadedDoc.ProtectionType != ProtectionType.ReadOnly)
            throw new InvalidOperationException("The document is not protected as expected.");
    }
}
