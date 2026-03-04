using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the existing DOC document.
        Document doc = new Document("input.doc");

        // Protect the document with read‑only protection and a password.
        doc.Protect(ProtectionType.ReadOnly, "MyPassword");

        // Save the protected document in DOC format.
        DocSaveOptions saveOptions = new DocSaveOptions(SaveFormat.Doc);
        doc.Save("output_protected.doc", saveOptions);
    }
}
