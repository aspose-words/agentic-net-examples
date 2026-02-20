using System;
using Aspose.Words;
using Aspose.Words.Saving;

class WriteProtectionExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add some content.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello world! This document is write‑protected.");

        // Set a write‑protection password (max 15 characters) and recommend read‑only.
        doc.WriteProtection.SetPassword("MyPassword");
        doc.WriteProtection.ReadOnlyRecommended = true;

        // Save the protected document as DOCX.
        // The path can be changed as needed.
        string outputPath = "WriteProtectedDocument.docx";
        doc.Save(outputPath, SaveFormat.Docx);

        // Optional: Verify that the document is write‑protected.
        Console.WriteLine($"IsWriteProtected: {doc.WriteProtection.IsWriteProtected}");
        Console.WriteLine($"Password valid: {doc.WriteProtection.ValidatePassword("MyPassword")}");
    }
}
