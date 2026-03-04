using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;

class ProtectDocumentExample
{
    static void Main()
    {
        // Path where the document will be saved.
        string artifactsDir = @"C:\Temp\"; // Adjust as needed.
        Directory.CreateDirectory(artifactsDir);

        // Create a new blank document.
        Document doc = new Document();

        // Add some content to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This document is protected with a password.");

        // Apply read‑only protection and set a password.
        // Users opening the file in Microsoft Word will be prompted for this password.
        doc.Protect(ProtectionType.ReadOnly, "MyPassword");

        // Save the protected document in DOCX format.
        string protectedPath = Path.Combine(artifactsDir, "ProtectedDocument.docx");
        doc.Save(protectedPath);

        // -----------------------------------------------------------------
        // The following demonstrates how to load the protected document
        // using the same password (optional, shown for completeness).
        // -----------------------------------------------------------------
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.Password = "MyPassword"; // Supply the password for opening.
        Document loadedDoc = new Document(protectedPath, loadOptions);

        // The document can still be edited programmatically without providing the password.
        DocumentBuilder loadedBuilder = new DocumentBuilder(loadedDoc);
        loadedBuilder.Writeln("Additional text added programmatically.");

        // Save the modified document (still protected with the same password).
        string modifiedPath = Path.Combine(artifactsDir, "ProtectedDocument_Modified.docx");
        loadedDoc.Save(modifiedPath);
    }
}
