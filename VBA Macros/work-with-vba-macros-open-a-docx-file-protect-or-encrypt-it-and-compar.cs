using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Settings;

class VbaMacroProtectionAndComparison
{
    static void Main()
    {
        // Paths to the source documents.
        const string sourcePath1 = @"C:\Docs\DocumentWithMacros.docx";
        const string sourcePath2 = @"C:\Docs\AnotherDocument.docx";

        // -----------------------------------------------------------------
        // 1. Load the first document and check if it contains VBA macros.
        // -----------------------------------------------------------------
        Document doc1 = new Document(sourcePath1);

        // Detect file format information without fully loading the document.
        FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(sourcePath1);
        bool hasMacros = formatInfo.HasMacros;
        Console.WriteLine($"Document '{sourcePath1}' contains macros: {hasMacros}");

        // -----------------------------------------------------------------
        // 2. Protect the first document (read‑only) with a password.
        // -----------------------------------------------------------------
        const string protectPassword = "Protect123";
        doc1.Protect(ProtectionType.ReadOnly, protectPassword);

        // Save the protected document.
        const string protectedPath = @"C:\Docs\DocumentWithMacros_Protected.docx";
        doc1.Save(protectedPath);
        Console.WriteLine($"Protected document saved to: {protectedPath}");

        // -----------------------------------------------------------------
        // 3. Load the second document and encrypt it with a password.
        // -----------------------------------------------------------------
        Document doc2 = new Document(sourcePath2);

        // Use OoxmlSaveOptions to apply encryption when saving as .docx.
        const string encryptPassword = "Encrypt456";
        OoxmlSaveOptions encryptOptions = new OoxmlSaveOptions
        {
            Password = encryptPassword
        };

        const string encryptedPath = @"C:\Docs\AnotherDocument_Encrypted.docx";
        doc2.Save(encryptedPath, encryptOptions);
        Console.WriteLine($"Encrypted document saved to: {encryptedPath}");

        // -----------------------------------------------------------------
        // 4. Compare the original (unprotected) versions of the two documents.
        // -----------------------------------------------------------------
        // Reload the original documents to ensure we compare the unprotected content.
        Document originalDoc1 = new Document(sourcePath1);
        Document originalDoc2 = new Document(sourcePath2);

        // Perform the comparison; the result will be stored in originalDoc1.
        const string author = "ComparisonEngine";
        originalDoc1.Compare(originalDoc2, author, DateTime.Now);

        const string comparisonResultPath = @"C:\Docs\ComparisonResult.docx";
        originalDoc1.Save(comparisonResultPath);
        Console.WriteLine($"Comparison document saved to: {comparisonResultPath}");

        // -----------------------------------------------------------------
        // 5. (Optional) Verify that the comparison document contains revisions.
        // -----------------------------------------------------------------
        Document comparisonDoc = new Document(comparisonResultPath);
        int revisionCount = comparisonDoc.Revisions.Count;
        Console.WriteLine($"Number of revisions in comparison document: {revisionCount}");
    }
}
