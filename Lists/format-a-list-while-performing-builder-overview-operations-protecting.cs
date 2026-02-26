using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Loading; // Added for LoadOptions

namespace AsposeWordsDemo
{
    class Program
    {
        static void Main()
        {
            // 1. Create a new blank document.
            Document originalDoc = new Document();

            // 2. Use DocumentBuilder to format a numbered list.
            DocumentBuilder builder = new DocumentBuilder(originalDoc);
            builder.ListFormat.ApplyNumberDefault();          // Start a default numbered list.
            builder.Writeln("First item");                    // List item 1.
            builder.Writeln("Second item");                   // List item 2.
            builder.ListFormat.RemoveNumbers();               // End the list.

            // 3. Save the original document (lifecycle: create → save).
            originalDoc.Save("Original.docx");

            // 4. Clone the original document to create an edited version.
            Document editedDoc = (Document)originalDoc.Clone(true);

            // 5. Modify the edited document – add another list item.
            DocumentBuilder editBuilder = new DocumentBuilder(editedDoc);
            editBuilder.MoveToDocumentEnd();                  // Move cursor to the end.
            editBuilder.Writeln("Third item");                // New list item.

            // 6. Compare the original document with the edited one.
            //    The comparison will generate revision marks in the original document.
            originalDoc.Compare(editedDoc, "Comparer", DateTime.Now);

            // 7. Save the document that now contains the comparison revisions.
            originalDoc.Save("Compared.docx");

            // 8. Protect the compared document with a password (read‑only protection).
            originalDoc.Protect(ProtectionType.ReadOnly, "pwd123");

            // 9. Save the protected (encrypted) document.
            originalDoc.Save("Protected.docx");

            // 10. Demonstrate loading an encrypted document using LoadOptions.
            LoadOptions loadOptions = new LoadOptions { Password = "pwd123" }; // Correct usage of LoadOptions
            Document loadedProtected = new Document("Protected.docx", loadOptions);

            // 11. Verify that the loaded document can be accessed (e.g., get its text).
            Console.WriteLine("Loaded protected document text:");
            Console.WriteLine(loadedProtected.GetText().Trim());
        }
    }
}
