using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Paths to the files used in the example.
        string sourcePath = "Input.docx";          // Document to protect/encrypt.
        string otherPath = "Other.docx";           // Document to compare with.
        string protectedPath = "Protected.docx";   // Output of the protected document.
        string encryptedPath = "Encrypted.docx";   // Output of the encrypted document.
        string compareResultPath = "Comparison.docx"; // Output of the comparison result.

        // Load the source document.
        Document sourceDoc = new Document(sourcePath);

        // -------------------------------------------------
        // Protect the document (read‑only) with a password.
        // -------------------------------------------------
        sourceDoc.Protect(ProtectionType.ReadOnly, "protectPwd");
        // Save the protected version.
        sourceDoc.Save(protectedPath);

        // -------------------------------------------------
        // Encrypt the document using OOXML encryption.
        // -------------------------------------------------
        OoxmlSaveOptions encryptOptions = new OoxmlSaveOptions();
        encryptOptions.Password = "encryptPwd";
        // Save the encrypted version.
        sourceDoc.Save(encryptedPath, encryptOptions);

        // Load the document that will be compared against the source.
        Document otherDoc = new Document(otherPath);

        // -------------------------------------------------
        // Compare the two documents.
        // The result document will contain revisions that represent the differences.
        // -------------------------------------------------
        // Clone the source document because Compare mutates the document it is called on.
        Document comparison = (Document)sourceDoc.Clone();
        comparison.Compare(otherDoc, "Comparer", DateTime.Now);

        // Save the comparison result.
        comparison.Save(compareResultPath);
    }
}
