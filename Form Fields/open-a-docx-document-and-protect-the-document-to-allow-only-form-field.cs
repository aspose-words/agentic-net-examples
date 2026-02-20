using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load an existing DOCX file.
        Document doc = new Document("InputDocument.docx");

        // Protect the document so that only form fields can be edited.
        doc.Protect(ProtectionType.AllowOnlyFormFields);

        // Save the protected document.
        doc.Save("ProtectedDocument.docx");
    }
}
