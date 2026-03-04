using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Insert some sample text.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This document is protected with a password.");

        // Apply read‑only protection and set a password.
        // Users opening the file in Microsoft Word will be prompted for this password.
        doc.Protect(ProtectionType.ReadOnly, "MyPassword");

        // Save the protected document in DOCX format.
        doc.Save("ProtectedDocument.docx");
    }
}
