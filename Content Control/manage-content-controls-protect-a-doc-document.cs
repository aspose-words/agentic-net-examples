using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load an existing Word document.
        Document doc = new Document("input.docx");

        // Apply read‑only protection with a password.
        // This prevents editing in Microsoft Word unless the password is supplied.
        doc.Protect(ProtectionType.ReadOnly, "MyPassword");

        // Save the protected document in the legacy DOC format.
        // DocSaveOptions allows us to specify the target format explicitly.
        DocSaveOptions saveOptions = new DocSaveOptions(SaveFormat.Doc);
        doc.Save("output.doc", saveOptions);
    }
}
