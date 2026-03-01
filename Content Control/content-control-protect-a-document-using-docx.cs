using System;
using Aspose.Words;
using Aspose.Words.Settings;

class ProtectDocumentExample
{
    static void Main()
    {
        // Create a new blank Word document.
        Document doc = new Document();

        // Insert some sample text.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This document is protected with read‑only protection.");

        // Apply read‑only protection and set a password.
        doc.Protect(ProtectionType.ReadOnly, "SecretPwd");

        // Save the protected document in DOCX format.
        doc.Save("ProtectedDocument.docx");
    }
}
