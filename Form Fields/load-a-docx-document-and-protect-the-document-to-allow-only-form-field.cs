using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load an existing DOCX file from disk.
        Document doc = new Document("input.docx");

        // Protect the document so that users can only fill in form fields.
        doc.Protect(ProtectionType.AllowOnlyFormFields);

        // Save the protected document to a new file.
        doc.Save("output_protected.docx");
    }
}
