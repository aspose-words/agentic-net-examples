using Aspose.Words;
using Aspose.Words.Settings;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Insert some sample text.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This document is protected. Only comments can be edited.");

        // Apply document protection that allows only comments, with a password.
        doc.Protect(ProtectionType.AllowOnlyComments, "SecretPwd");

        // Set write‑protection settings (optional).
        doc.WriteProtection.SetPassword("WritePwd");
        doc.WriteProtection.ReadOnlyRecommended = true;

        // Save the protected document.
        doc.Save("ProtectedDocument.docx");
    }
}
