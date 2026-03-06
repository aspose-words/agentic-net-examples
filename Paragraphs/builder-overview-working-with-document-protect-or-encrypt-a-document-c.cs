using System;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

namespace AsposeWordsDemo
{
    class Program
    {
        static void Main()
        {
            // 1. Create a new blank document.
            Document originalDoc = new Document();

            // 2. Use DocumentBuilder to add some content.
            DocumentBuilder builder = new DocumentBuilder(originalDoc);
            builder.Writeln("This is the original document.");
            builder.Writeln("It contains two paragraphs for demonstration.");

            // 3. Save the original document.
            originalDoc.Save("Original.docx");

            // 4. Protect the document with a password (read‑only protection).
            originalDoc.Protect(ProtectionType.ReadOnly, "MyPassword");

            // 5. Save the protected document.
            originalDoc.Save("Protected.docx");

            // 6. Load the protected document using the password.
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.Password = "MyPassword";
            Document protectedDoc = new Document("Protected.docx", loadOptions);

            // 7. Unprotect the document (remove protection).
            protectedDoc.Unprotect("MyPassword");

            // 8. Save the unprotected version.
            protectedDoc.Save("Unprotected.docx");

            // 9. Create an edited version of the original document.
            Document editedDoc = new Document("Original.docx");
            DocumentBuilder editBuilder = new DocumentBuilder(editedDoc);
            // Move to the end of the document and add a new paragraph.
            editBuilder.MoveToDocumentEnd();
            editBuilder.Writeln("This paragraph was added in the edited version.");

            // 10. Save the edited document.
            editedDoc.Save("Edited.docx");

            // 11. Compare the original and edited documents.
            // The comparison will produce revisions in the original document.
            originalDoc.Compare(editedDoc, "Comparer", DateTime.Now);

            // 12. Save the document that now contains revision marks.
            originalDoc.Save("Compared.docx");

            // Optional: Accept all revisions to transform the original into the edited version.
            originalDoc.Revisions.AcceptAll();
            originalDoc.Save("AcceptedRevisions.docx");
        }
    }
}
