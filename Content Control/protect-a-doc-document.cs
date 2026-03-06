using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new document.
        Document doc = new Document();

        // Insert some text.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This document is protected.");

        // Apply read‑only protection with a password.
        doc.Protect(ProtectionType.ReadOnly, "SecretPwd");

        // Save the document in the legacy DOC format.
        DocSaveOptions saveOptions = new DocSaveOptions(SaveFormat.Doc);
        // Uncomment the next line to set a password that protects opening the file.
        // saveOptions.Password = "OpenPwd";

        doc.Save("ProtectedDocument.doc", saveOptions);
    }
}
