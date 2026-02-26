using System;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Paths to the source documents.
        const string docPath1 = @"C:\Docs\Original.docx";
        const string docPath2 = @"C:\Docs\Modified.docx";

        // -----------------------------------------------------------------
        // 1. Load the first document (no password required).
        // -----------------------------------------------------------------
        Document doc1 = new Document(docPath1);

        // -----------------------------------------------------------------
        // 2. Protect the document (read‑only) and then encrypt it with a password.
        //    Protection limits editing in Word, while encryption prevents opening
        //    without the password.
        // -----------------------------------------------------------------
        doc1.Protect(ProtectionType.ReadOnly, "protectPwd");

        // Use OoxmlSaveOptions to set an encryption password.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
        {
            Password = "encryptPwd"
        };
        const string encryptedPath = @"C:\Docs\Original_Protected_Encrypted.docx";
        doc1.Save(encryptedPath, saveOptions);

        // -----------------------------------------------------------------
        // 3. Load the encrypted document using the correct password.
        // -----------------------------------------------------------------
        LoadOptions loadOptions = new LoadOptions("encryptPwd");
        Document protectedDoc = new Document(encryptedPath, loadOptions);

        // -----------------------------------------------------------------
        // 4. Load the second document that we want to compare against.
        // -----------------------------------------------------------------
        Document doc2 = new Document(docPath2);

        // -----------------------------------------------------------------
        // 5. Compare the two documents.
        //    The result is stored in the 'protectedDoc' as revision changes.
        // -----------------------------------------------------------------
        // The author name and the comparison date are optional.
        protectedDoc.Compare(doc2, "Comparer", DateTime.Now);

        // -----------------------------------------------------------------
        // 6. Save the comparison result (includes revisions) to a new file.
        // -----------------------------------------------------------------
        const string comparisonResultPath = @"C:\Docs\ComparisonResult.docx";
        protectedDoc.Save(comparisonResultPath);
    }
}
