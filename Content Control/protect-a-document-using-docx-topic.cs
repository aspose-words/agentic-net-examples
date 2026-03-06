using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ProtectDocumentExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add some content to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This document will be protected with a password.");

        // Protect the document so that it is read‑only.
        // The second argument sets the password that Microsoft Word will require.
        doc.Protect(ProtectionType.ReadOnly, "MyPassword");

        // Save the protected document in DOCX format.
        doc.Save("ProtectedDocument.docx", SaveFormat.Docx);
    }
}
