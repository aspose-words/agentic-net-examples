using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ProtectDocumentExample
{
    static void Main()
    {
        // Define the folder where the output files will be saved.
        // Adjust this path as needed for your environment.
        string artifactsDir = @"C:\Temp\Aspose\Docs\";

        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add some content.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello world! This document is write‑protected.");

        // Set a write‑protection password (max 15 characters) and recommend read‑only.
        doc.WriteProtection.SetPassword("MyPassword");
        doc.WriteProtection.ReadOnlyRecommended = true;

        // Save the protected document as DOCX.
        doc.Save(artifactsDir + "WriteProtected.docx");

        // Load the saved document to verify the protection settings.
        Document loadedDoc = new Document(artifactsDir + "WriteProtected.docx");

        // The document can still be edited programmatically.
        DocumentBuilder loadedBuilder = new DocumentBuilder(loadedDoc);
        loadedBuilder.MoveToDocumentEnd();
        loadedBuilder.Writeln("Additional text added programmatically.");

        // Output verification results.
        Console.WriteLine("Is write protected: " + loadedDoc.WriteProtection.IsWriteProtected);
        Console.WriteLine("Password valid: " + loadedDoc.WriteProtection.ValidatePassword("MyPassword"));
        Console.WriteLine("Document text:");
        Console.WriteLine(loadedDoc.GetText());

        // Optionally, save the document again after programmatic changes.
        loadedDoc.Save(artifactsDir + "WriteProtected_Modified.docx");
    }
}
