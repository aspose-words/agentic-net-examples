using System;
using Aspose.Words;

namespace AsposeWordsExample
{
    class Program
    {
        static void Main()
        {
            // Load an existing DOCX document from the file system.
            // The Document(string) constructor follows the provided load rule.
            Document doc = new Document("InputDocument.docx");

            // Protect the document so that only form fields can be edited.
            // This uses the Protect(ProtectionType) method as defined in the rules.
            doc.Protect(ProtectionType.AllowOnlyFormFields);

            // Save the protected document back to the file system.
            // The Save(string) method follows the provided save rule.
            doc.Save("ProtectedDocument.docx");
        }
    }
}
