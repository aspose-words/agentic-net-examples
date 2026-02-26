using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // List to keep a log of builder operations.
        List<string> log = new List<string>();

        // 1. Create a new blank document using DocumentBuilder's default constructor.
        DocumentBuilder builder = new DocumentBuilder();
        log.Add("Created blank document with DocumentBuilder.");

        // 2. Insert a paragraph into the document.
        builder.Writeln("This is the original document.");
        log.Add("Inserted paragraph into document.");

        // 3. Save the original document to disk.
        string originalPath = "Original.docx";
        builder.Document.Save(originalPath);
        log.Add($"Saved original document to '{originalPath}'.");

        // 4. Load the saved document from file.
        Document originalDoc = new Document(originalPath);
        log.Add("Loaded original document from file.");

        // 5. Apply read‑only protection with a password.
        originalDoc.Protect(ProtectionType.ReadOnly, "SecretPwd");
        log.Add("Applied read‑only protection with password.");

        // 6. Save the protected document with encryption (password on save).
        string protectedPath = "Protected.docx";
        OoxmlSaveOptions protectOptions = new OoxmlSaveOptions(SaveFormat.Docx)
        {
            Password = "SecretPwd"
        };
        originalDoc.Save(protectedPath, protectOptions);
        log.Add($"Saved protected (encrypted) document to '{protectedPath}'.");

        // 7. Create a second document that will be compared against the original.
        DocumentBuilder builder2 = new DocumentBuilder();
        builder2.Writeln("This is the edited document with extra line.");
        builder2.Writeln("Additional content.");
        string editedPath = "Edited.docx";
        builder2.Document.Save(editedPath);
        log.Add($"Created and saved edited document to '{editedPath}'.");

        // 8. Load the edited document.
        Document editedDoc = new Document(editedPath);
        log.Add("Loaded edited document from file.");

        // 9. Compare the original (protected) document with the edited document.
        // Revisions will be added to the original document.
        originalDoc.Compare(editedDoc, "Comparer", DateTime.Now);
        log.Add("Compared original document with edited document, revisions added.");

        // 10. Save the comparison result.
        string compareResultPath = "ComparisonResult.docx";
        originalDoc.Save(compareResultPath);
        log.Add($"Saved comparison result to '{compareResultPath}'.");

        // Output the operation log.
        foreach (string entry in log)
        {
            Console.WriteLine(entry);
        }
    }
}
