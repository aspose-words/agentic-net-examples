using System;
using Aspose.Words;
using Aspose.Words.Settings;

class ProtectDocumentExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add some content using DocumentBuilder.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This document will be protected.");

        // Protect the document for read‑only editing and set a password.
        // The password is required only when the document is opened in Microsoft Word.
        doc.Protect(ProtectionType.ReadOnly, "MySecretPwd");

        // Additionally, set write‑protection (read‑only recommendation) with a password.
        // This does not encrypt the file; it only prevents accidental edits.
        doc.WriteProtection.SetPassword("WritePwd");
        doc.WriteProtection.ReadOnlyRecommended = true;

        // Save the protected document to a DOCX file.
        doc.Save("ProtectedDocument.docx");
    }
}
