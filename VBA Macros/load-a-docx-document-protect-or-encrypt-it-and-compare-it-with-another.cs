using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Paths to the source documents and output files.
        string sourcePath1 = "Input1.docx";
        string sourcePath2 = "Input2.docx";
        string protectedPath = "Protected.docx";
        string comparisonResultPath = "ComparisonResult.docx";

        // Load the first document.
        Document doc1 = new Document(sourcePath1);

        // Apply read‑only protection with a password.
        doc1.Protect(ProtectionType.ReadOnly, "protectPwd");

        // Save the protected document with encryption (password on save).
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
        saveOptions.Password = "encryptPwd";
        doc1.Save(protectedPath, saveOptions);

        // Load the second document (assumed not encrypted).
        Document doc2 = new Document(sourcePath2);

        // Compare the two documents. In older Aspose.Words versions Compare returns void,
        // so we invoke it directly on doc1 and then save the same document which now contains revision marks.
        doc1.Compare(doc2, "Comparer", DateTime.Now);

        // Save the comparison result.
        doc1.Save(comparisonResultPath);
    }
}
