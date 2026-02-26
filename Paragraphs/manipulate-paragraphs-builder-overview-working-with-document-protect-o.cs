using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // -------------------------------------------------
        // 1. Create a new document and add a few paragraphs
        // -------------------------------------------------
        Document doc1 = new Document();                     // create blank document
        DocumentBuilder builder = new DocumentBuilder(doc1);
        builder.Writeln("First paragraph in Document 1.");
        builder.Writeln("Second paragraph in Document 1.");

        // -------------------------------------------------
        // 2. Protect the document (read‑only) with a password
        // -------------------------------------------------
        doc1.Protect(ProtectionType.ReadOnly, "protectPwd");

        // Save the protected document
        doc1.Save("ProtectedDoc1.docx");

        // -------------------------------------------------
        // 3. Encrypt the same document with a password
        // -------------------------------------------------
        OoxmlSaveOptions encryptOptions = new OoxmlSaveOptions(SaveFormat.Docx)
        {
            Password = "encryptPwd"
        };
        doc1.Save("EncryptedDoc1.docx", encryptOptions);

        // -------------------------------------------------
        // 4. Load a second document for comparison
        // -------------------------------------------------
        // For demonstration we load the protected version we just saved.
        Document doc2 = new Document("ProtectedDoc1.docx");

        // -------------------------------------------------
        // 5. Prepare an original (unprotected) version to compare against
        // -------------------------------------------------
        Document original = new Document();
        DocumentBuilder origBuilder = new DocumentBuilder(original);
        origBuilder.Writeln("First paragraph in Document 1.");
        // Change the second paragraph to create a difference
        origBuilder.Writeln("Second paragraph in Document 1 - modified.");

        // -------------------------------------------------
        // 6. Compare the two documents – revisions are stored in the result
        // -------------------------------------------------
        // The Compare method modifies the calling document (original) and does not return a value in older API versions.
        // Therefore we call it without assignment and then use the modified 'original' as the result.
        original.Compare(doc2, "Comparer", DateTime.Now);
        Document comparisonResult = original; // now contains revision marks

        // Save the comparison result (shows insert/delete revisions)
        comparisonResult.Save("ComparisonResult.docx");
    }
}
