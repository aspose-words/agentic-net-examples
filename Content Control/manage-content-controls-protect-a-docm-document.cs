using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Settings;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add some content.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a DOCM document that will be protected.");

        // Set write‑protection password and recommend read‑only opening.
        doc.WriteProtection.SetPassword("Secret123");
        doc.WriteProtection.ReadOnlyRecommended = true;

        // Apply document protection (read‑only) with the same password.
        // This prevents editing unless the correct password is supplied.
        doc.Protect(ProtectionType.ReadOnly, "Secret123");

        // Save the document as a macro‑enabled DOCM file.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docm);
        doc.Save("ProtectedDocument.docm", saveOptions);
    }
}
